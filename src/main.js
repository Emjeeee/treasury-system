import "./style.css";
import "./storage.js";
import * as XLSX from "xlsx";
import Chart from "chart.js/auto";

let _ExcelJS = null;
async function loadExcelJS() {
  if (!_ExcelJS) _ExcelJS = (await import("exceljs")).default;
  return _ExcelJS;
}
let _GridStack = null;
// editor drag/resize dashboard cuma dipakai admin, sekali-sekali - jangan
// muat gridstack (+ CSS-nya) di bundle utama yg selalu diunduh semua orang
async function loadGridstack() {
  if (!_GridStack) {
    await import("gridstack/dist/gridstack.min.css");
    _GridStack = (await import("gridstack")).GridStack;
  }
  return _GridStack;
}

/* ================= penyimpanan ================= */
// prefix isolasi untuk verifikasi/testing, jangan pernah sentuh key produksi
const KEY_PREFIX = import.meta.env.VITE_KEY_PREFIX || "";
const SKEY = KEY_PREFIX + "ct_shared_v4",
  HKEY = KEY_PREFIX + "ct_events_index_v1",
  LKEY = KEY_PREFIX + "ct_session_v4";
const eventKey = (id) => KEY_PREFIX + "ct_event_" + id + "_v1";
// (*) = kolom wajib diisi saat impor, lihat parseRows()
const DEFAULT_TPL = [
  { f: "date", h: "Tanggal*" },
  { f: "name", h: "Nama*" },
  { f: "phone", h: "Nomor WhatsApp" },
  { f: "cat", h: "Kategori*" },
  { f: "note", h: "Keterangan" },
  { f: "note2", h: "Keterangan Tambahan" },
  { f: "amount", h: "Nominal*" },
  { f: "bank", h: "Metode Pembayaran*" },
  { f: "bankName", h: "Bank" },
  { f: "status", h: "Status*" },
];
// id statis (bukan uid()) supaya bisa dipanggil saat modul dimuat, sebelum
// uid() sendiri didefinisikan lebih bawah di file ini
const DEFAULT_METHODS = () => [
  { id: "m_bank", name: "Transfer Bank", type: "bank" },
  { id: "m_tunai", name: "Tunai", type: "cash" },
  { id: "m_cek", name: "Cek/Giro", type: "cheque" },
  { id: "m_utang", name: "Utang", type: "debt" },
];
// daftar bank umum di Indonesia - dipakai sbg field tambahan "Bank" pas
// metode pembayaran yg dipilih bertipe "bank" (lihat onMethodChange())
const INDONESIA_BANKS = [
  "BCA", "Bank Mandiri", "BRI", "BNI", "CIMB Niaga", "Bank Danamon", "Bank Permata",
  "BTN", "BSI (Bank Syariah Indonesia)", "Bank Muamalat", "OCBC NISP", "Panin Bank",
  "Maybank Indonesia", "UOB Indonesia", "HSBC Indonesia", "Bank Jago", "SeaBank",
  "Jenius (BTPN)", "Bank Mega", "Bank Sinarmas", "Bank BJB", "Bank DKI", "Bank Jateng",
  "Bank Jatim", "Bank Sumut", "Bank Nagari", "Bank Riau Kepri", "Bank Sumsel Babel",
  "Bank Kalbar", "Bank Kaltimtara", "Bank Sulselbar", "Bank NTB Syariah", "Bank Bali",
  "Commonwealth Bank", "Standard Chartered", "Citibank", "DBS Indonesia", "ANZ Indonesia",
  "Bank Woori Saudara", "Bank Sahabat Sampoerna", "Bank Neo Commerce", "Allo Bank",
  "Bank Index", "Bank Ganesha", "Bank Victoria", "Lainnya",
];
// ubah config.banks (daftar nama string, versi lama) jadi config.methods
// (daftar {id,name,type}) - dipanggil begitu acara dimuat, sekali per acara
function methodsFromBanks(banks) {
  return (banks || []).map((name) => ({
    id: uid(),
    name,
    type: /tunai|cash/i.test(name) ? "cash" : "bank",
  }));
}
// S = data acara yang sedang dibuka (config+tx+logs). G = data lintas-acara
// (daftar acara, daftar pengguna) di baris kv_store terpisah (HKEY).
let S = {
  rev: 0,
  config: {
    event: "Konser Amal 2026",
    date: "",
    cats: [
      { id: "c_reg", n: "Reguler", group: "income", hasQty: true, p: 150000, q: 300 },
      { id: "c_vip", n: "VIP", group: "income", hasQty: true, p: 400000, q: 80 },
    ],
    methods: DEFAULT_METHODS(),
    tpl: DEFAULT_TPL,
  },
  tx: [],
  logs: [],
};
let G = { rev: 0, events: [], users: [], logs: [] };
let currentEventId = null;
let me = null,
  imp = null,
  lang = "id",
  theme = "light";
let screen = "boot",
  tab = "dash",
  filter = "all",
  panel = "seats",
  adminTab = "set",
  hubTab = "events",
  authMode = "in";
let chartInstances = {},
  dashGrid = null,
  pendingImp = [],
  logQ = "",
  txQ = "",
  txSort = { k: "date", dir: "desc" },
  // default "all" spy angka laporan yg diunduh cocok dgn KPI dashboard
  // (yg selalu all-time, tidak per-periode) - admin yg memang mau laporan
  // per-minggu/bulan masih bisa pilih manual di dialog export
  exportKind = "all",
  monitorCache = null;
let txPage = 1,
  logPage = 1,
  boardBuyPage = 1,
  boardDonPage = 1,
  hubEventsPage = 1,
  hubEventsQ = "",
  hubEventsFilter = "all",
  hubEventsSort = { k: "name", dir: "asc" },
  hubStaffPage = 1,
  hubStaffQ = "",
  hubStaffFilter = "all",
  hubStaffSort = { k: "name", dir: "asc" };
const PAGE_SIZE = 20;

const D = () => S.config;
const rp = (n) => "Rp " + (Math.round(n) || 0).toLocaleString("id-ID");
const rpk = (n) => {
  const a = Math.abs(n);
  return a >= 1e9
    ? "Rp " + (n / 1e9).toFixed(1) + " M"
    : a >= 1e6
      ? "Rp " + (n / 1e6).toFixed(1) + (lang === "id" ? " jt" : " M")
      : rp(n);
};
const uid = () =>
  Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const esc = (s) =>
  String(s ?? "").replace(
    /[&<>"]/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c],
  );
// kategori bertingkat (mis. "Pembelian Kursi" + tier "Platinum") ditampilkan
// gabungan dimana pun kategori transaksi muncul di UI/laporan
const catLabel = (x) => (x.tier ? `${x.cat} - ${x.tier}` : x.cat || "");
const today = () => new Date().toISOString().slice(0, 10);
const narrow = () => window.innerWidth < 960;
const now = () => new Date().toISOString();
const acting = () => imp || me; // identitas yang sedang dipakai
const isAdmin = () => me && me.role === "admin" && !imp; // hak admin hilang saat menyamar
// role "viewer": cuma bisa lihat data (dashboard), tidak bisa
// tambah/ubah/hapus apa pun - ikut identitas yg SEDANG aktif (termasuk saat
// admin menyamar sbg viewer), beda dgn isAdmin() yg sengaja ikut identitas asli
const canEdit = () => acting()?.role !== "viewer";
const roleLabel = (role) =>
  role === "admin" ? t("admins") : role === "viewer" ? t("viewer") : t("treas");
const roleTagClass = (role) =>
  role === "admin" ? "t-adm" : role === "viewer" ? "t-view" : "t-tic";
// potong array jadi satu halaman (dipakai semua tampilan list: transaksi,
// log, peringkat, acara, staff) supaya tidak perlu scroll ratusan baris
function paginate(arr, page, size = PAGE_SIZE) {
  const totalPages = Math.max(1, Math.ceil(arr.length / size));
  const p = Math.min(Math.max(1, page), totalPages);
  return { items: arr.slice((p - 1) * size, p * size), page: p, totalPages, total: arr.length };
}
function pagerHtml(page, totalPages, setter) {
  if (totalPages <= 1) return "";
  return `<div class="rowsp pager" style="flex:none;justify-content:center;gap:20px">
    <button class="btn ghost sm icon" ${page <= 1 ? "disabled" : ""} onclick="${setter}(${page - 1})" aria-label="prev">‹</button>
    <span class="hint mono" style="min-width:52px;text-align:center">${page} / ${totalPages}</span>
    <button class="btn ghost sm icon" ${page >= totalPages ? "disabled" : ""} onclick="${setter}(${page + 1})" aria-label="next">›</button>
  </div>`;
}

async function sha(s) {
  const b = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode("kas." + s),
  );
  return [...new Uint8Array(b)]
    .map((x) => x.toString(16).padStart(2, "0"))
    .join("");
}
function logEntry(act, det) {
  return {
    ts: now(),
    user: acting().name,
    email: acting().email,
    role: acting().role,
    asAdmin: imp ? me.email : "",
    act,
    det: det || "",
  };
}
async function pullEvent() {
  if (!currentEventId) return false;
  try {
    const r = await window.storage.get(eventKey(currentEventId), true);
    if (r && r.value) {
      const v = JSON.parse(r.value);
      if (v.rev !== S.rev) {
        S = Object.assign(S, v);
        return true;
      }
    }
  } catch (e) {
    console.error("pullEvent() failed:", e);
  }
  return false;
}
async function pushEvent() {
  if (!currentEventId) return;
  S.rev = (S.rev || 0) + 1;
  try {
    await window.storage.set(eventKey(currentEventId), JSON.stringify(S), true);
  } catch (e) {
    console.error("pushEvent() failed:", e);
    toast("⚠ " + t("saveErr"));
  }
}
async function mutateEvent(fn, logAct, logDet) {
  try {
    const r = await window.storage.get(eventKey(currentEventId), true);
    if (r && r.value) S = Object.assign(S, JSON.parse(r.value));
  } catch (e) {
    console.error("mutateEvent() pre-fetch failed:", e);
  }
  fn();
  if (logAct) {
    S.logs.unshift(logEntry(logAct, logDet));
    S.logs = S.logs.slice(0, 600);
  }
  await pushEvent();
}
// pull/push/mutate tetap dipakai di seluruh kode yang mengubah data acara
// yang sedang dibuka (transaksi, pengaturan, log acara itu) - tidak perlu
// diganti satu per satu, cukup alias ke versi *Event di atas.
const pull = pullEvent,
  push = pushEvent,
  mutate = mutateEvent;
async function pullHub() {
  try {
    const r = await window.storage.get(HKEY, true);
    if (r && r.value) {
      const v = JSON.parse(r.value);
      if (v.rev !== G.rev) {
        G = Object.assign(G, v);
        return true;
      }
    }
  } catch (e) {
    console.error("pullHub() failed:", e);
  }
  return false;
}
async function pushHub() {
  G.rev = (G.rev || 0) + 1;
  try {
    await window.storage.set(HKEY, JSON.stringify(G), true);
  } catch (e) {
    console.error("pushHub() failed:", e);
    toast("⚠ " + t("saveErr"));
  }
}
async function mutateHub(fn, logAct, logDet) {
  try {
    const r = await window.storage.get(HKEY, true);
    if (r && r.value) G = Object.assign(G, JSON.parse(r.value));
  } catch (e) {
    console.error("mutateHub() pre-fetch failed:", e);
  }
  fn();
  if (logAct) {
    G.logs.unshift(logEntry(logAct, logDet));
    G.logs = G.logs.slice(0, 600);
  }
  await pushHub();
}
async function saveSession() {
  try {
    await window.storage.set(
      LKEY,
      JSON.stringify({
        email: me ? me.email : "",
        imp: imp ? imp.email : "",
        lang,
        theme,
        eventId: currentEventId,
      }),
      false,
    );
  } catch (e) {}
}
function downloadJSON(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
async function ensureMigrated() {
  const idx = await window.storage.get(HKEY, true);
  if (idx && idx.value) return; // sudah dimigrasi (atau instalasi baru multi-acara)

  const old = await window.storage.get(SKEY, true);
  if (!old || !old.value) {
    // instalasi baru, belum ada data lama sama sekali
    await window.storage.set(
      HKEY,
      JSON.stringify({ rev: 1, events: [], users: [], logs: [] }),
      true,
    );
    return;
  }

  const legacy = JSON.parse(old.value);
  // 1. backup sisi server (jaring pengaman utama)
  await window.storage.set(
    KEY_PREFIX + "ct_backup_pre_migration_v1",
    JSON.stringify({ backedUpAt: now(), source: SKEY, data: legacy }),
    true,
  );
  // 1b. unduhan lokal (bonus, bisa diblokir browser mobile, bukan jaring pengaman utama)
  try {
    downloadJSON(legacy, `Backup-PreMigration-${today()}.json`);
  } catch (e) {}

  // 2. id acara deterministik: kalau ada 2 client migrasi bersamaan, hasilnya sama
  const EVID = "ev_default";
  await window.storage.set(
    eventKey(EVID),
    JSON.stringify({
      rev: 1,
      config: {
        event: legacy.config.event,
        date: legacy.config.date,
        cats: legacy.config.cats,
        methods: methodsFromBanks(legacy.config.banks),
        tpl: legacy.config.tpl,
      },
      tx: legacy.tx || [],
      logs: legacy.logs || [],
    }),
    true,
  );

  // 3. baris index ditulis TERAKHIR - keberadaannya adalah penanda "sudah dimigrasi"
  const users = (legacy.users || []).map((u) => ({
    ...u,
    eventIds: u.role === "admin" ? [] : [EVID],
  }));
  await window.storage.set(
    HKEY,
    JSON.stringify({
      rev: 1,
      events: [
        {
          id: EVID,
          name: legacy.config.event,
          date: legacy.config.date,
          status: "active",
          createdAt: now(),
          createdBy: "system-migration",
        },
      ],
      users,
      logs: [],
    }),
    true,
  );
  // baris lama ct_shared_v4 sengaja dibiarkan apa adanya sebagai cadangan tambahan
}
// buat baris acara baru yang masih kosong (dipakai saat setup akun admin
// pertama, dan nanti oleh Konsol Admin untuk menambah acara baru)
async function createEventRow(id, name, date, createdBy) {
  await window.storage.set(
    eventKey(id),
    JSON.stringify({
      rev: 1,
      config: {
        event: name,
        date: date || "",
        cats: [],
        methods: [],
        tpl: DEFAULT_TPL,
        dashboard: { widgets: starterDashboardWidgets() },
      },
      tx: [],
      logs: [],
    }),
    true,
  );
  return { id, name, date: date || "", status: "active", createdAt: now(), createdBy };
}
const closeIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>`;
const trendUpIcon = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M4 17l6-6 4 4 6-8"/><path d="M14 7h6v6"/></svg>`;
const trendDownIcon = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7l6 6 4-4 6 8"/><path d="M14 17h6v-6"/></svg>`;
const closeBtn = () =>
  `<button type="button" class="btn ghost sm icon" onclick="closeSheet()" aria-label="${t("close")}" title="${t("close")}">${closeIcon}</button>`;
const proofIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M21 16l-5-4-4 3-3-2-6 5"/></svg>`;
const downloadIcon = `<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>`;
function toast(m) {
  const e = document.createElement("div");
  e.className = "toast";
  e.textContent = m;
  document.body.appendChild(e);
  setTimeout(() => e.remove(), 2400);
}

/* ================= bahasa ================= */
const L = {
  id: {
    dash: "Dashboard",
    tx: "Transaksi",
    board: "Peringkat",
    admin: "Admin",
    set: "Pengaturan",
    users: "Pengguna",
    logs: "Log aktivitas",
    xls: "Unduh Laporan",
    add: "+ Transaksi",
    net: "Saldo bersih",
    ticket: "Penjualan kursi",
    don: "Donasi & investor",
    incomeW: "Penerimaan",
    exp: "Pengeluaran",
    wait: "Belum lunas",
    trend: "Penerimaan {n} hari terakhir",
    seats: "Ketersediaan kursi",
    left: "kursi tersisa",
    soldOf: "{a} terjual dari {b}",
    queue: "Belum lunas",
    noQueue: "Semua transaksi sudah lunas.",
    topBuy: "Pembeli kursi terbanyak",
    topDon: "Donatur & investor terbesar",
    inRek: "Lunas",
    edit: "Ubah",
    all: "Semua",
    unchecked: "Belum dicek",
    seatsW: "Kursi",
    expW: "Pengeluaran",
    date: "Tanggal",
    name: "Nama",
    detail: "Rincian",
    amount: "Jumlah",
    noTx: "Belum ada transaksi.",
    noTxSub: "Tekan + Transaksi untuk mencatat yang pertama.",
    newTx: "Catat transaksi",
    editTx: "Ubah transaksi",
    txDetail: "Detail transaksi",
    close: "Tutup",
    paymentMethod: "Metode pembayaran",
    bankNameLbl: "Bank",
    nameGeneric: "Nama",
    namePay: "Dibayarkan kepada",
    tierLbl: "Tipe",
    searchPlaceholder: "Cari...",
    wa: "Nomor WhatsApp",
    txCat: "Kategori",
    noCatYet: "Belum ada kategori. Tambahkan dulu di Pengaturan.",
    nSeat: "Jumlah kursi",
    amtIn: "Nominal transfer",
    amtOut: "Nominal pengeluaran",
    amtCash: "Jumlah tunai",
    amtCheque: "Nominal cek",
    amtDebt: "Nominal utang",
    note: "Catatan bukti transfer",
    note2: "Keterangan tambahan",
    proof: "Bukti transfer (gambar)",
    proofBig: "Ukuran gambar maksimal 8 MB",
    proofHint: "Foto otomatis dikecilkan agar hemat penyimpanan",
    chooseFile: "Pilih berkas",
    noFileChosen: "Tidak ada berkas dipilih",
    fileAttached: "Gambar terpasang",
    viewProof: "Lihat bukti",
    searchTx: "Cari nama, telepon, catatan, rekening...",
    status: "Status",
    payStatus: "Status pembayaran",
    stW: "Belum lunas",
    stV: "Lunas",
    save: "Simpan transaksi",
    saveGeneric: "Simpan",
    del: "Hapus",
    saved: "Transaksi tersimpan",
    deleted: "Transaksi dihapus",
    verified: "Ditandai sudah masuk rekening",
    event: "Acara",
    evName: "Nama acara",
    evDate: "Tanggal acara",
    methods: "Metode pembayaran",
    methodName: "Nama",
    methodType: "Tipe",
    methodBank: "Transfer Bank",
    methodCash: "Tunai",
    methodCheque: "Cek",
    methodOther: "Lainnya",
    methodDebt: "Utang",
    addMethod: "+ Tambah metode",
    chequeNo: "No. cek",
    chequeBank: "Bank penerbit",
    chequeDate: "Tanggal cair",
    saveSet: "Simpan pengaturan",
    setSaved: "Pengaturan disimpan",
    catTitle: "Kategori transaksi",
    price: "Harga",
    quota: "Kuota",
    groupLbl: "Kelompok",
    qtyLbl: "Lacak jumlah",
    tiersFor: "Tingkat harga: {n}",
    tiersHint: "Tambahkan tingkat/tipe dengan harga berbeda (mis. Platinum/Gold/Silver). Kategori dengan tingkat akan menampilkan pilihan tipe di form transaksi.",
    tierName: "Nama tingkat",
    addTier: "+ Tambah tingkat",
    tiersBtn: "Tingkat ({n})",
    bonusRulesTitle: "Bonus otomatis",
    bonusRulesHint: "Contoh: sponsor minimal Rp 20 juta berhak dapat 2 kursi Platinum gratis. Ini hanya pengingat di form transaksi - transaksi kursi gratisnya tetap dicatat manual.",
    bonusMinAmt: "Nominal minimal",
    bonusTargetCat: "Kategori hadiah",
    bonusTargetTier: "Tingkat",
    bonusQty: "Jumlah gratis",
    addBonusRule: "+ Tambah aturan bonus",
    bonusNeedsTiers: "Buat kategori bertingkat dulu (mis. Pembelian Kursi) sebelum menambah aturan bonus.",
    bonusHint: "Berhak mendapat {qty} {cat} tingkat {tier} gratis. Catat sebagai transaksi terpisah dengan nominal Rp 0.",
    addCat: "+ Tambah kategori",
    dataT: "Data",
    dataP: "Data tersimpan bersama untuk semua pengguna aplikasi ini.",
    clearAll: "Hapus semua transaksi",
    cleared: "Semua transaksi dihapus",
    imp: "Unggah Data",
    impDrop: "Pilih berkas Excel atau CSV",
    impSub: "Gunakan template agar kolom terbaca otomatis",
    dlTpl: "Unduh template",
    tplTitle: "Kolom template",
    tplP: "Ubah nama kolom bila berkas Anda memakai istilah lain.",
    field: "Data",
    colName: "Nama kolom di berkas",
    prev: "Pratinjau impor",
    rowsOk: "{n} baris siap diimpor",
    txCount: "{n} transaksi",
    drillSearchPlaceholder: "Cari nama, kategori, nominal, atau status...",
    txCountLbl: "transaksi",
    totalLbl: "total",
    changeLbl: "Perubahan",
    dailyBalanceHint: "Saldo kumulatif per hari, dibandingkan hari sebelumnya (H-1).",
    vsYesterday: "vs kemarin",
    rowsBad: "{n} baris dilewati",
    doImp: "Impor {n} baris",
    imported: "{n} transaksi diimpor",
    noCol: "Kolom tidak dikenali. Periksa nama kolom di Pengaturan.",
    trans: "transaksi",
    free: "gratis",
    seatsSold: "kursi terjual",
    noneYet: "Belum ada data.",
    confirmDel: "Hapus transaksi ini?",
    confirmAll: "Hapus semua transaksi? Tindakan ini tidak bisa dibatalkan.",
    rank: "Penyumbang terbesar",
    signIn: "Masuk",
    signUp: "Daftar",
    signOut: "Keluar",
    email: "Email",
    pass: "Kata sandi",
    pass2: "Ulangi kata sandi",
    fullName: "Nama lengkap",
    google: "Lanjut dengan Google",
    forgot: "Lupa kata sandi?",
    haveAcc: "Sudah punya akun? Masuk",
    noAcc: "Belum punya akun? Daftar",
    welcome: "Treasury System",
    welcomeSub: "Sistem pencatatan keuangan acara.",
    setupT: "Buat akun admin pertama",
    setupSub: "Akun ini mengelola pengaturan, pengguna, dan log aktivitas.",
    createAdmin: "Buat akun admin",
    role: "Peran",
    admins: "Admin",
    treas: "Treasurer",
    viewer: "Viewer",
    viewOnlyNotice: "Anda masuk sebagai viewer - hanya dapat melihat data, tidak dapat menambah/mengubah/menghapus.",
    badLogin: "Email atau kata sandi salah.",
    emailUsed: "Email sudah terdaftar.",
    passShort: "Kata sandi minimal 8 karakter.",
    passDiff: "Kata sandi tidak sama.",
    fillAll: "Semua kolom wajib diisi.",
    hello: "Selamat datang, {n}",
    addUser: "+ Tambah pengguna",
    userAdded: "Pengguna ditambahkan",
    userSaved: "Perubahan pengguna disimpan",
    userDel: "Pengguna dihapus",
    confirmUser: "Hapus pengguna ini?",
    lastAdmin: "Minimal harus ada satu admin aktif.",
    active: "Aktif",
    inactive: "Nonaktif",
    resetPass: "Atur ulang kata sandi",
    newPass: "Kata sandi baru",
    passReset: "Kata sandi diperbarui",
    loginAs: "Masuk sebagai",
    impBanner: "Anda melihat sebagai {n}",
    backAdmin: "Kembali ke admin",
    joined: "Bergabung",
    lastIn: "Terakhir masuk",
    never: "Belum pernah",
    forgotT: "Atur ulang kata sandi",
    forgotSub:
      "Masukkan email dan kode pemulihan yang Anda buat saat mendaftar.",
    recov: "Kode pemulihan",
    recovNew: "Kode pemulihan (simpan baik-baik)",
    recovHint: "Dipakai bila lupa kata sandi. Admin juga bisa mengatur ulang.",
    badRecov: "Email atau kode pemulihan salah.",
    reqAdmin: "Atau minta admin mengatur ulang kata sandi Anda.",
    actLogin: "Masuk",
    actLogout: "Keluar",
    actCreate: "Tambah transaksi",
    actUpdate: "Ubah transaksi",
    actDelete: "Hapus transaksi",
    actVerify: "Verifikasi mutasi",
    actImport: "Impor data",
    actExport: "Ekspor Excel",
    actSet: "Ubah pengaturan",
    actUser: "Kelola pengguna",
    actImp: "Masuk sebagai pengguna",
    actClear: "Hapus semua transaksi",
    actSignup: "Daftar akun",
    actEventCreate: "Buat acara",
    actEventArchive: "Ubah status acara",
    actEventDuplicate: "Duplikat acara",
    searchLog: "Cari nama, aksi, atau rincian",
    time: "Waktu",
    user: "Pengguna",
    action: "Aksi",
    info: "Rincian",
    exportLog: "Unduh log",
    noLog: "Belum ada aktivitas.",
    onlyAdmin: "Hanya admin yang dapat membuka halaman ini.",
    syncOn: "Data diperbarui",
    saveErr: "Gagal menyimpan",
    gTitle: "Masuk dengan Google",
    gSub: "Demo tanpa server: masukkan alamat Gmail Anda untuk membuat atau membuka akun.",
    gGo: "Lanjutkan",
    byAdmin: "oleh admin {n}",
    bulk: "Aksi massal",
    themeToggle: "Ganti tema terang/gelap",
    copyright: "Hak cipta dilindungi.",
    customCol: "Kustom",
    addCol: "+ Tambah kolom kustom",
    customFields: "Kolom tambahan",
    perWeek: "Mingguan",
    perMonth: "Bulanan",
    perAll: "Semua data",
    perDate: "Tanggal acuan",
    perAllLabel: "Semua periode",
    adminConsole: "Konsol Admin",
    events: "Acara",
    staff: "Staff",
    monitoring: "Monitoring",
    newEvent: "+ Acara baru",
    archive: "Arsipkan",
    restore: "Aktifkan lagi",
    duplicate: "Duplikat",
    duplicateEvent: "Duplikat acara",
    duplicateEventHint: "Membuat acara baru dengan kategori, metode pembayaran, dan tata letak dashboard yang sama seperti \"{n}\". Transaksi dan log TIDAK ikut disalin.",
    archived: "Diarsipkan",
    chooseEvent: "Pilih acara",
    chooseEventSub: "Anda dipetakan ke beberapa acara. Pilih salah satu untuk mulai.",
    switchEvent: "Ganti acara",
    noEventsAssigned: "Belum ada acara yang ditugaskan ke Anda. Hubungi admin.",
    needsAttention: "Perlu perhatian",
    staffPerformance: "Performa staff",
    enterWorkspace: "Buka",
    backToConsole: "Kembali ke Konsol Admin",
    financialSummary: "Ringkasan keuangan",
    recorded: "Dicatat",
    refresh: "Segarkan",
    verified: "Diverifikasi",
    dashboardTab: "Dashboard",
    editDashboard: "Atur dashboard",
    dashEditorHint: "Geser & ubah ukuran widget, lalu simpan tata letaknya.",
    addWidget: "+ Tambah widget",
    saveLayout: "Simpan tata letak",
    widgetType: "Jenis widget",
    widgetTitleLbl: "Judul widget",
    metricLbl: "Angka yang ditampilkan",
    catFilterLbl: "Batasi ke kategori",
    catFilterHint: "Kosongkan (tidak ada yang dipilih) untuk memakai semua kategori di kelompok ini.",
    wKpi: "Kartu angka",
    wChart: "Grafik tren",
    wPie: "Grafik pai",
    wTable: "Tabel transaksi",
    wBreakdown: "Rincian per kategori",
    wQuota: "Kuota / ketersediaan",
    wQueue: "Antrean menunggu",
    wRank: "Papan peringkat",
    chartStyleLbl: "Tampilan grafik",
    chartStyleBar: "Batang",
    chartStyleLine: "Garis",
    allTypesLbl: "Semua",
    ofTotal: "dari total",
    clickForDetail: "Klik untuk lihat detail",
    tableMoreHint: "+{n} transaksi lain, buka penuh di tab Transaksi",
  },
  en: {
    dash: "Dashboard",
    tx: "Transactions",
    board: "Rankings",
    admin: "Admin",
    set: "Settings",
    users: "Users",
    logs: "Activity log",
    xls: "Download Report",
    add: "+ Transaction",
    net: "Net balance",
    ticket: "Seat sales",
    don: "Donations & investors",
    incomeW: "Income",
    exp: "Expenses",
    wait: "Unpaid",
    trend: "Income, last {n} days",
    seats: "Seat availability",
    left: "seats left",
    soldOf: "{a} sold of {b}",
    queue: "Unpaid",
    noQueue: "Everything is paid off.",
    topBuy: "Top seat buyers",
    topDon: "Top donors & investors",
    inRek: "Paid",
    edit: "Edit",
    all: "All",
    unchecked: "Unchecked",
    seatsW: "Seats",
    expW: "Expense",
    date: "Date",
    name: "Name",
    detail: "Details",
    amount: "Amount",
    noTx: "No transactions yet.",
    noTxSub: "Tap + Transaction to record the first one.",
    newTx: "Record transaction",
    editTx: "Edit transaction",
    txDetail: "Transaction detail",
    close: "Close",
    paymentMethod: "Payment method",
    bankNameLbl: "Bank",
    nameGeneric: "Name",
    namePay: "Paid to",
    tierLbl: "Type",
    searchPlaceholder: "Search...",
    wa: "WhatsApp number",
    txCat: "Category",
    noCatYet: "No categories yet. Add one in Settings first.",
    nSeat: "Number of seats",
    amtIn: "Transfer amount",
    amtOut: "Expense amount",
    amtCash: "Cash amount",
    amtCheque: "Cheque amount",
    amtDebt: "Debt amount",
    note: "Transfer proof note",
    note2: "Additional note",
    proof: "Transfer proof (image)",
    proofBig: "Image must be under 8 MB",
    proofHint: "Photos are automatically resized to save space",
    chooseFile: "Choose file",
    noFileChosen: "No file chosen",
    fileAttached: "Image attached",
    viewProof: "View proof",
    searchTx: "Search name, phone, note, account...",
    status: "Status",
    payStatus: "Payment status",
    stW: "Unpaid",
    stV: "Paid",
    save: "Save transaction",
    saveGeneric: "Save",
    del: "Delete",
    saved: "Transaction saved",
    deleted: "Transaction deleted",
    verified: "Marked as received",
    event: "Event",
    evName: "Event name",
    evDate: "Event date",
    methods: "Payment methods",
    methodName: "Name",
    methodType: "Type",
    methodBank: "Bank Transfer",
    methodCash: "Cash",
    methodCheque: "Cheque",
    methodOther: "Other",
    methodDebt: "Debt",
    addMethod: "+ Add method",
    chequeNo: "Cheque no.",
    chequeBank: "Issuing bank",
    chequeDate: "Clearing date",
    saveSet: "Save settings",
    setSaved: "Settings saved",
    catTitle: "Transaction categories",
    price: "Price",
    quota: "Quota",
    groupLbl: "Group",
    qtyLbl: "Track qty",
    tiersFor: "Price tiers: {n}",
    tiersHint: "Add tiers/types with different prices (e.g. Platinum/Gold/Silver). Categories with tiers show a type picker on the transaction form.",
    tierName: "Tier name",
    addTier: "+ Add tier",
    tiersBtn: "Tiers ({n})",
    bonusRulesTitle: "Automatic bonus",
    bonusRulesHint: "Example: sponsors giving at least Rp 20 million get 2 free Platinum seats. This is just a reminder on the transaction form - the free-seat transaction still needs to be recorded manually.",
    bonusMinAmt: "Minimum amount",
    bonusTargetCat: "Reward category",
    bonusTargetTier: "Tier",
    bonusQty: "Free quantity",
    addBonusRule: "+ Add bonus rule",
    bonusNeedsTiers: "Create a tiered category first (e.g. Seat Purchase) before adding a bonus rule.",
    bonusHint: "Entitled to {qty} free {tier}-tier {cat}. Record it as a separate Rp 0 transaction.",
    addCat: "+ Add category",
    dataT: "Data",
    dataP: "Data is shared across everyone using this app.",
    clearAll: "Delete all transactions",
    cleared: "All transactions deleted",
    imp: "Upload Data",
    impDrop: "Choose an Excel or CSV file",
    impSub: "Use the template so columns are detected automatically",
    dlTpl: "Download template",
    tplTitle: "Template columns",
    tplP: "Rename a column if your file uses different wording.",
    field: "Data",
    colName: "Column name in file",
    prev: "Import preview",
    rowsOk: "{n} rows ready to import",
    txCount: "{n} transactions",
    drillSearchPlaceholder: "Search name, category, amount, or status...",
    txCountLbl: "transactions",
    totalLbl: "total",
    changeLbl: "Change",
    dailyBalanceHint: "Cumulative daily balance, compared to the previous day.",
    vsYesterday: "vs yesterday",
    rowsBad: "{n} rows skipped",
    doImp: "Import {n} rows",
    imported: "{n} transactions imported",
    noCol: "Columns not recognised. Check column names in Settings.",
    trans: "transactions",
    free: "free",
    seatsSold: "seats sold",
    noneYet: "No data yet.",
    confirmDel: "Delete this transaction?",
    confirmAll: "Delete all transactions? This cannot be undone.",
    rank: "Biggest contributors",
    signIn: "Sign in",
    signUp: "Sign up",
    signOut: "Sign out",
    email: "Email",
    pass: "Password",
    pass2: "Repeat password",
    fullName: "Full name",
    google: "Continue with Google",
    forgot: "Forgot password?",
    haveAcc: "Already have an account? Sign in",
    noAcc: "No account yet? Sign up",
    welcome: "Treasury System",
    welcomeSub: "Event financial recording system.",
    setupT: "Create the first admin account",
    setupSub: "This account manages settings, users and the activity log.",
    createAdmin: "Create admin account",
    role: "Role",
    admins: "Admin",
    treas: "Treasurer",
    viewer: "Viewer",
    viewOnlyNotice: "You're signed in as a viewer - you can only view data, not add/edit/delete it.",
    badLogin: "Wrong email or password.",
    emailUsed: "That email is already registered.",
    passShort: "Password must be at least 8 characters.",
    passDiff: "Passwords do not match.",
    fillAll: "All fields are required.",
    hello: "Welcome, {n}",
    addUser: "+ Add user",
    userAdded: "User added",
    userSaved: "User updated",
    userDel: "User deleted",
    confirmUser: "Delete this user?",
    lastAdmin: "At least one active admin is required.",
    active: "Active",
    inactive: "Inactive",
    resetPass: "Reset password",
    newPass: "New password",
    passReset: "Password updated",
    loginAs: "Sign in as",
    impBanner: "Viewing as {n}",
    backAdmin: "Back to admin",
    joined: "Joined",
    lastIn: "Last sign-in",
    never: "Never",
    forgotT: "Reset password",
    forgotSub: "Enter your email and the recovery code you saved at sign-up.",
    recov: "Recovery code",
    recovNew: "Recovery code (keep it safe)",
    recovHint: "Used if you forget your password. An admin can also reset it.",
    badRecov: "Wrong email or recovery code.",
    reqAdmin: "Or ask an admin to reset your password.",
    actLogin: "Signed in",
    actLogout: "Signed out",
    actCreate: "Created transaction",
    actUpdate: "Updated transaction",
    actDelete: "Deleted transaction",
    actVerify: "Verified payment",
    actImport: "Imported data",
    actExport: "Exported Excel",
    actSet: "Changed settings",
    actUser: "Managed users",
    actImp: "Signed in as user",
    actClear: "Deleted all transactions",
    actSignup: "Created account",
    actEventCreate: "Created event",
    actEventArchive: "Changed event status",
    actEventDuplicate: "Duplicated event",
    searchLog: "Search name, action or details",
    time: "Time",
    user: "User",
    action: "Action",
    info: "Details",
    exportLog: "Export log",
    noLog: "No activity yet.",
    onlyAdmin: "Only admins can open this page.",
    syncOn: "Data updated",
    saveErr: "Could not save",
    gTitle: "Sign in with Google",
    gSub: "Server-free demo: enter your Gmail address to create or open an account.",
    gGo: "Continue",
    byAdmin: "by admin {n}",
    bulk: "Bulk actions",
    themeToggle: "Toggle light/dark theme",
    copyright: "All rights reserved.",
    customCol: "Custom",
    addCol: "+ Add custom column",
    customFields: "Additional fields",
    perWeek: "Weekly",
    perMonth: "Monthly",
    perAll: "All data",
    perDate: "Reference date",
    perAllLabel: "All time",
    adminConsole: "Admin Console",
    events: "Events",
    staff: "Staff",
    monitoring: "Monitoring",
    newEvent: "+ New event",
    archive: "Archive",
    restore: "Restore",
    duplicate: "Duplicate",
    duplicateEvent: "Duplicate event",
    duplicateEventHint: "Creates a new event with the same categories, payment methods, and dashboard layout as \"{n}\". Transactions and logs are NOT copied.",
    archived: "Archived",
    chooseEvent: "Choose an event",
    chooseEventSub: "You're mapped to more than one event. Pick one to get started.",
    switchEvent: "Switch event",
    noEventsAssigned: "No events assigned to you yet. Contact your admin.",
    needsAttention: "Needs attention",
    staffPerformance: "Staff performance",
    enterWorkspace: "Open",
    backToConsole: "Back to Admin Console",
    financialSummary: "Financial summary",
    recorded: "Recorded",
    refresh: "Refresh",
    verified: "Verified",
    dashboardTab: "Dashboard",
    editDashboard: "Edit dashboard",
    dashEditorHint: "Drag & resize widgets, then save the layout.",
    addWidget: "+ Add widget",
    saveLayout: "Save layout",
    widgetType: "Widget type",
    widgetTitleLbl: "Widget title",
    metricLbl: "Value shown",
    catFilterLbl: "Limit to categories",
    catFilterHint: "Leave empty (none selected) to use every category in this group.",
    wKpi: "Number card",
    wChart: "Trend chart",
    wPie: "Pie chart",
    wTable: "Transaction table",
    wBreakdown: "Category breakdown",
    wQuota: "Quota / availability",
    wQueue: "Pending queue",
    wRank: "Leaderboard",
    chartStyleLbl: "Chart style",
    chartStyleBar: "Bar",
    chartStyleLine: "Line",
    allTypesLbl: "All",
    ofTotal: "of total",
    clickForDetail: "Click to view detail",
    tableMoreHint: "+{n} more, open the Transactions tab to see all",
  },
};
const t = (k, v = {}) =>
  String(L[lang][k] || k).replace(/\{(\w+)\}/g, (_, x) => v[x]);
const FL = {
  no: { id: "No", en: "No" },
  date: { id: "Tanggal", en: "Date" },
  type: { id: "Jenis", en: "Type" },
  name: { id: "Nama", en: "Name" },
  phone: { id: "No. WhatsApp", en: "WhatsApp No." },
  cat: { id: "Kategori", en: "Category" },
  seats: { id: "Jumlah Kursi", en: "Seat Count" },
  debit: { id: "Debit (Rp)", en: "Debit (Rp)" },
  credit: { id: "Kredit (Rp)", en: "Credit (Rp)" },
  amount: { id: "Nominal", en: "Amount" },
  bank: { id: "Metode Pembayaran", en: "Payment Method" },
  bankName: { id: "Bank", en: "Bank" },
  status: { id: "Status", en: "Status" },
  note: { id: "Catatan", en: "Note" },
  note2: { id: "Keterangan Tambahan", en: "Additional Note" },
};

/* ================= hitung ================= */
function sums(tx = S.tx, cats = D().cats) {
  const v = tx.filter((x) => x.status === "verified");
  const income = v
    .filter((x) => x.type === "income")
    .reduce((a, b) => a + b.amount, 0);
  const exp = v
    .filter((x) => x.type === "expense")
    .reduce((a, b) => a + b.amount, 0);
  // kuota/ketersediaan cuma berlaku utk kategori income yg "melacak jumlah"
  // (dulu implisit = semua kategori tiket, sekarang eksplisit per kategori)
  const qtyCatNames = cats
    .filter((c) => c.group === "income" && c.hasQty)
    .map((c) => c.n);
  const pd = tx.filter((x) => x.status === "pending" && x.type !== "expense");
  const quota = cats
    .filter((c) => c.group === "income" && c.hasQty)
    .reduce((a, c) => a + (+c.q || 0), 0);
  const sold = v
    .filter((x) => x.type === "income" && qtyCatNames.includes(x.cat))
    .reduce((a, b) => a + (+b.seats || 0), 0);
  const held = pd
    .filter((x) => qtyCatNames.includes(x.cat))
    .reduce((a, b) => a + (+b.seats || 0), 0);
  const byCat = {};
  v.forEach((x) => {
    byCat[x.cat] = (byCat[x.cat] || 0) + x.amount;
  });
  return {
    income,
    exp,
    quota,
    sold,
    held,
    byCat,
    pendAmt: pd.reduce((a, b) => a + b.amount, 0),
    pendN: pd.length,
    net: income - exp,
    avail: Math.max(quota - sold - held, 0),
  };
}
function seriesByDay(d) {
  const o = [],
    m = {};
  S.tx
    .filter((x) => x.status === "verified" && x.type !== "expense")
    .forEach((x) => {
      m[x.date] = (m[x.date] || 0) + x.amount;
    });
  for (let i = d - 1; i >= 0; i--) {
    const dt = new Date();
    dt.setDate(dt.getDate() - i);
    const k = dt.toISOString().slice(0, 10);
    o.push({ k, label: k.slice(8) + "/" + k.slice(5, 7), v: m[k] || 0 });
  }
  return o;
}
// scope: undefined = semua income; array nama kategori = batasi ke kategori itu
// (dipakai buat pisah papan peringkat "kategori bertipe qty" vs "bukan", lihat vBoard/pRank)
function byPerson(scope, tx = S.tx) {
  const m = {};
  tx
    .filter(
      (x) =>
        x.status === "verified" &&
        x.type === "income" &&
        (!scope || scope.includes(x.cat)),
    )
    .forEach((x) => {
      const k = (x.name || "—").trim();
      if (!m[k]) m[k] = { name: k, amount: 0, seats: 0, n: 0 };
      m[k].amount += x.amount;
      m[k].seats += +x.seats || 0;
      m[k].n++;
    });
  return Object.values(m).sort((a, b) => b.amount - a.amount);
}

/* ================= boot & sinkronisasi ================= */
function myEvents(user) {
  if (!user) return [];
  return user.role === "admin"
    ? G.events.filter((e) => e.status === "active")
    : G.events.filter(
        (e) => e.status === "active" && (user.eventIds || []).includes(e.id),
      );
}
async function enterEvent(id) {
  currentEventId = id;
  // rev:0 tak pernah dipakai baris sungguhan (mulai dari 1), jadi pullEvent()
  // di bawah selalu menimpa S dengan data acara yang baru, tidak pernah
  // "dikira sama" dan dilewati gara-gara kebetulan rev acara lama == acara baru
  S = {
    rev: 0,
    config: { event: "", date: "", cats: [], methods: [], tpl: DEFAULT_TPL },
    tx: [],
    logs: [],
  };
  await pullEvent();
  await normalizeEventConfig();
}
// acara lama (sebelum config.methods ada) masih pakai config.banks (array
// nama string) - ubah ke methods sekali saat acara itu pertama kali dibuka
// lagi, lalu simpan supaya tidak perlu dikonversi ulang tiap kali
async function normalizeEventConfig() {
  if (!S.config.methods) {
    const methods = methodsFromBanks(S.config.banks);
    await mutateEvent(() => {
      S.config.methods = methods;
      delete S.config.banks;
    }, null);
  }
  if (needsCatMigration()) {
    await mutateEvent(() => migrateCatsAndTypes(), null);
  }
}
// v3.0: dulu kategori cuma tier kursi {n,p,q}, transaksi punya 3 tipe tetap
// (ticket/donation - keduanya berbagi daftar kategori yg sama - dan expense,
// yg kategorinya freetext "Keperluan"). Sekarang tiap kategori eksplisit
// punya group (income/expense) + hasQty (melacak jumlah&harga satuan, atau
// nominal manual), dan tipe transaksi cuma income/expense (ikut grup
// kategori yg dipilih). Migrasi sekali per acara, idempoten - dicek lewat
// bentuk data yg sudah ada, bukan version counter.
function needsCatMigration() {
  return (
    (S.config.cats || []).some((c) => !c.group) ||
    (S.tx || []).some((x) => x.type === "ticket" || x.type === "donation") ||
    !S.config.dashboard
  );
}
function migrateCatsAndTypes() {
  S.config.cats.forEach((c) => {
    if (!c.group) {
      c.id = c.id || uid();
      c.group = "income";
      c.hasQty = true;
    }
  });
  // baris expense lama nyimpan "Keperluan" bebas di x.cat - jadikan kategori
  // expense baru per nilai unik, supaya datanya tetap tampil terkelompok
  // dgn benar (admin bisa gabung/ganti nama lagi lewat Pengaturan)
  const expNames = [
    ...new Set(
      S.tx
        .filter((x) => x.type === "expense" && x.cat)
        .map((x) => String(x.cat).trim())
        .filter(Boolean),
    ),
  ];
  expNames.forEach((n) => {
    if (!S.config.cats.some((c) => c.group === "expense" && c.n === n)) {
      S.config.cats.push({
        id: uid(),
        n,
        group: "expense",
        hasQty: false,
        p: 0,
        q: 0,
      });
    }
  });
  S.tx.forEach((x) => {
    if (x.type === "ticket" || x.type === "donation") x.type = "income";
  });
  if (!S.config.dashboard) {
    S.config.dashboard = { widgets: defaultDashboardWidgets(S.config.cats) };
  }
}
// keputusan layar mana yang dibuka setelah login/impersonate/reload:
// - acara tersimpan di sesi masih valid -> lanjutkan di situ
// - staff dengan tepat 1 acara -> langsung masuk, tanpa perlu memilih
// - admin tanpa acara tersimpan -> Konsol Admin
// - selain itu (staff 0 atau 2+ acara) -> layar pilih acara
function resolveEventEntry(user, savedEventId) {
  const avail = myEvents(user);
  const saved = avail.find((e) => e.id === savedEventId);
  if (saved) return { screen: "app", eventId: saved.id };
  // admin selalu mendarat di Konsol Admin dulu (kecuali sesi tersimpan di atas
  // sudah mengarah ke satu acara) - beda dari staff yang auto-masuk kalau
  // acaranya cuma satu
  if (user.role === "admin") return { screen: "hub", eventId: null };
  if (avail.length === 1) return { screen: "app", eventId: avail[0].id };
  return { screen: "picker", eventId: null };
}
async function enterResolved(user, savedEventId) {
  const r = resolveEventEntry(user, savedEventId);
  if (r.eventId) await enterEvent(r.eventId);
  else currentEventId = null;
  screen = r.screen;
}
async function boot() {
  await ensureMigrated();
  await pullHub();
  let savedEventId = null;
  try {
    const r = await window.storage.get(LKEY, false);
    if (r && r.value) {
      const v = JSON.parse(r.value);
      lang = v.lang || "id";
      theme = v.theme || "light";
      me = G.users.find((u) => u.email === v.email && u.active) || null;
      if (me && v.imp) imp = G.users.find((u) => u.email === v.imp) || null;
      savedEventId = v.eventId || null;
    }
  } catch (e) {}
  applyTheme();
  if (!G.users.length) {
    screen = "setup";
  } else if (!me) {
    screen = "auth";
  } else {
    await enterResolved(acting(), savedEventId);
  }
  render();
  setInterval(async () => {
    if (screen === "app") {
      const ch = await pullEvent();
      await pullHub();
      if (me) {
        const u = G.users.find((x) => x.email === me.email);
        if (!u || !u.active) {
          await doSignOut(true);
          return;
        }
        me = u;
        if (imp) imp = G.users.find((x) => x.email === imp.email) || null;
      }
      if (ch) render();
    } else if (screen === "picker" || screen === "hub") {
      const ch = await pullHub();
      if (ch) render();
    }
  }, 6000);
}
window.addEventListener("resize", () => {
  clearTimeout(window._r);
  window._r = setTimeout(() => {
    if (screen === "app" && tab === "dash") render();
  }, 180);
});

/* ================= render utama ================= */
function render() {
  const r = document.getElementById("root");
  if (screen === "setup") r.innerHTML = vSetup();
  else if (screen === "auth") r.innerHTML = vAuth();
  else if (screen === "picker") r.innerHTML = vPicker();
  else if (screen === "hub") r.innerHTML = vHub();
  else {
    r.innerHTML = vApp();
    // .fit (dashboard) mengurangi tinggi viewport dengan --banner supaya tidak
    // tabrakan dengan nav bawah saat banner "Viewing as" sedang tampil
    const bannerEl = r.querySelector(".banner");
    document.documentElement.style.setProperty(
      "--banner",
      (bannerEl ? bannerEl.offsetHeight : 0) + "px",
    );
    if (tab === "dash") {
      drawChart();
      drawPies();
      // render() bikin ulang #panelSeg dari nol tiap kali (klik tab, polling
      // 6dtk, dll) - scrollLeft-nya otomatis balik ke 0 walau tab aktifnya
      // (class "on") ada di posisi paling kanan, jadi tombol aktif jadi
      // tersembunyi dari layar mobile. Paksa scroll tombol aktif itu ke
      // dalam area yg terlihat setiap render, bukan cuma saat awal dibuka.
      document.getElementById("panelSeg")?.querySelector("button.on")
        ?.scrollIntoView({ block: "nearest", inline: "nearest" });
    } else if (tab === "admin" && adminTab === "dashboard") initDashGrid();
    else if (dashGrid) {
      dashGrid.destroy(false);
      dashGrid = null;
    }
  }
}
function langSeg() {
  return `<div class="seg" style="flex:none">${["id", "en"]
    .map(
      (l) =>
        `<button class="${lang === l ? "on" : ""}" onclick="setLang('${l}')">${l.toUpperCase()}</button>`,
    )
    .join("")}</div>`;
}
async function setLang(l) {
  lang = l;
  await saveSession();
  render();
}
const copyrightLine = () =>
  `<div class="hint" style="text-align:center;margin-top:16px">© ${new Date().getFullYear()} Michael Jonathan. ${t("copyright")}</div>`;
const sunIcon = `<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>`;
const moonIcon = `<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`;
const themeBtn = () =>
  `<button type="button" class="btn ghost sm icon" onclick="toggleTheme()" title="${t("themeToggle")}" aria-label="${t("themeToggle")}">${theme === "dark" ? sunIcon : moonIcon}</button>`;
function applyTheme() {
  document.documentElement.setAttribute("data-theme", theme);
}
async function toggleTheme() {
  theme = theme === "dark" ? "light" : "dark";
  applyTheme();
  await saveSession();
  render();
}

/* ================= layar setup & auth ================= */
function vSetup() {
  if (!setupRec) setupRec = recoveryCode();
  return `<div class="auth"><div class="authbox">
  <div class="rowsp" style="margin-bottom:14px;flex-wrap:wrap;row-gap:10px"><div style="display:flex;gap:10px;align-items:center;min-width:0;flex:1 1 auto">
    <img class="mark" src="/treasurySystem.ico" alt="Treasury System"><h1 style="font-size:19px">${t("setupT")}</h1></div>
    <div style="display:flex;gap:6px;align-items:center;flex:none">${themeBtn()}${langSeg()}</div></div>
  <div class="card" style="padding:20px">
    <p class="hint" style="margin:0 0 14px">${t("setupSub")}</p>
    <div class="field"><label>${t("fullName")}</label>
      <input id="a_name" placeholder="Michael Jonathan"></div>
    <div class="field"><label>${t("email")}</label>
      <input id="a_email" type="email" autocomplete="username" placeholder="michael@email.com"></div>
    <div class="field"><label>${t("pass")}</label>
      <div class="pw"><input id="a_pass" type="password" autocomplete="new-password" placeholder="Minimal 8 karakter">
        <button type="button" class="eye" onclick="toggleEye(this)">${eyeOn}</button></div></div>
    <div class="field"><label>${t("recovNew")}</label>
      <input id="a_rec" value="${setupRec}" readonly>
      <div class="hint" style="margin-top:4px">${t("recovHint")}</div></div>
    <button class="btn wide" onclick="createFirst()">${t("createAdmin")}</button>
  </div>${copyrightLine()}</div></div>`;
}
async function createFirst() {
  const rc = setupRec;
  const n = val("a_name"),
    e = val("a_email").toLowerCase(),
    p = val("a_pass");
    // rc = val(setupRec);
  if (!n || !e || !p) return toast(t("fillAll"));
  if (p.length < 8) return toast(t("passShort"));
  const u = {
    id: uid(),
    name: n,
    email: e,
    role: "admin",
    active: true,
    provider: "password",
    pass: await sha(e + p),
    rec: await sha(e + rc.toUpperCase()),
    created: now(),
    last: now(),
    eventIds: [],
  };
  me = u;
  imp = null;
  const eventEntry = await createEventRow(
    "ev_" + uid(),
    lang === "id" ? "Acara Pertama" : "First Event",
    "",
    e,
  );
  await mutateHub(() => {
    G.users.push(u);
    G.events.push(eventEntry);
  }, "actSignup", "admin " + e);
  currentEventId = eventEntry.id;
  await pullEvent();
  await saveSession();
  screen = "app";
  render();
  toast(t("hello", { n }));
}
const val = (i) => (document.getElementById(i)?.value || "").trim();

function vAuth() {
  const f = authMode;
  return `<div class="auth"><div class="authbox">
  <div class="rowsp" style="margin-bottom:14px;flex-wrap:wrap;row-gap:10px"><div style="display:flex;gap:10px;align-items:center;min-width:0;flex:1 1 auto">
    <img class="mark" src="/treasurySystem.ico" alt="Treasury System"><div style="min-width:0"><h1 style="font-size:19px">${f === "forgot" ? t("forgotT") : t("welcome")}</h1>
    <div class="hint">${f === "forgot" ? t("forgotSub") : t("welcomeSub")}</div></div></div>
    <div style="display:flex;gap:6px;align-items:center;flex:none">${themeBtn()}${langSeg()}</div></div>
  <div class="card" style="padding:20px">
  ${
    f === "in"
      ? `
    <div class="field"><label>${t("email")}</label><input id="i_email" type="email" autocomplete="username" placeholder="michael@email.com"></div>
    <div class="field"><label>${t("pass")}</label><input id="i_pass" type="password" autocomplete="current-password" placeholder="••••••••"
      onkeydown="if(event.key==='Enter')doSignIn()"></div>
    <button class="btn wide" onclick="doSignIn()">${t("signIn")}</button>
    <div class="divider">atau</div>
    <button class="gbtn" onclick="googleFlow()">${gIcon()} ${t("google")}</button>
    <div class="rowsp" style="margin-top:14px">
      <button class="btn ghost sm" onclick="setAuthMode('forgot')">${t("forgot")}</button>
      <button class="btn ghost sm" onclick="setAuthMode('up')">${t("signUp")}</button></div>`
      : f === "up"
        ? `
    <div class="field"><label>${t("fullName")}</label><input id="u_name"></div>
    <div class="field"><label>${t("email")}</label><input id="u_email" type="email" autocomplete="username"></div>
    <div class="field"><label>${t("pass")}</label><input id="u_pass" type="password" autocomplete="new-password"></div>
    <div class="field"><label>${t("pass2")}</label><input id="u_pass2" type="password" autocomplete="new-password"></div>
    <div class="field"><label>${t("recovNew")}</label><input id="u_rec" value="${uid().toUpperCase()}" readonly>
      <div class="hint" style="margin-top:4px">${t("recovHint")}</div></div>
    <button class="btn wide" onclick="doSignUp()">${t("signUp")}</button>
    <div class="divider">atau</div>
    <button class="gbtn" onclick="googleFlow()">${gIcon()} ${t("google")}</button>
    <button class="btn ghost sm wide" style="margin-top:14px" onclick="setAuthMode('in')">${t("haveAcc")}</button>`
        : `
    <div class="field"><label>${t("email")}</label><input id="r_email" type="email"></div>
    <div class="field"><label>${t("recov")}</label><input id="r_rec"></div>
    <div class="field"><label>${t("newPass")}</label><input id="r_pass" type="password"></div>
    <button class="btn wide" onclick="doForgot()">${t("resetPass")}</button>
    <p class="hint" style="margin:12px 0 0">${t("reqAdmin")}</p>
    <button class="btn ghost sm wide" style="margin-top:10px" onclick="setAuthMode('in')">${t("signIn")}</button>`
  }
  </div>${copyrightLine()}</div></div>`;
}
function gIcon() {
  return `<svg width="18" height="18" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.5 0 6.6 1.2 9 3.6l6.7-6.7C35.6 2.6 30.1 0 24 0 14.6 0 6.5 5.4 2.6 13.2l7.8 6.1C12.3 13.2 17.6 9.5 24 9.5z"/><path fill="#4285F4" d="M46.5 24.5c0-1.6-.1-3.2-.4-4.7H24v9h12.7c-.5 3-2.2 5.5-4.7 7.2l7.3 5.7c4.3-4 6.8-9.9 6.8-17.2z"/><path fill="#FBBC05" d="M10.4 28.7c-.5-1.5-.8-3-.8-4.7s.3-3.2.8-4.7l-7.8-6.1C.9 16.4 0 20.1 0 24s.9 7.6 2.6 10.8l7.8-6.1z"/><path fill="#34A853" d="M24 48c6.5 0 11.9-2.1 15.9-5.8l-7.3-5.7c-2 1.4-4.7 2.3-8.6 2.3-6.4 0-11.7-3.7-13.6-8.8l-7.8 6.1C6.5 42.6 14.6 48 24 48z"/></svg>`;
}

async function doSignIn() {
  const e = val("i_email").toLowerCase(),
    p = val("i_pass");
  await pullHub();
  const u = G.users.find((x) => x.email === e && x.active);
  if (!u || u.pass !== (await sha(e + p))) return toast(t("badLogin"));
  me = u;
  imp = null;
  await mutateHub(() => {
    G.users.find((v) => v.email === e).last = now();
  }, "actLogin", "");
  await enterResolved(me, null);
  await saveSession();
  tab = "dash";
  render();
  toast(t("hello", { n: u.name }));
}
async function doSignUp() {
  const n = val("u_name"),
    e = val("u_email").toLowerCase(),
    p = val("u_pass"),
    p2 = val("u_pass2"),
    rc = val("u_rec");
  if (!n || !e || !p || !rc) return toast(t("fillAll"));
  if (p.length < 8) return toast(t("passShort"));
  if (p !== p2) return toast(t("passDiff"));
  await pullHub();
  if (G.users.some((x) => x.email === e)) return toast(t("emailUsed"));
  const u = {
    id: uid(),
    name: n,
    email: e,
    role: "treasurer",
    active: true,
    provider: "password",
    pass: await sha(e + p),
    rec: await sha(e + rc.toUpperCase()),
    created: now(),
    last: now(),
    eventIds: [],
  };
  me = u;
  imp = null;
  await mutateHub(() => {
    G.users.push(u);
  }, "actSignup", e);
  await saveSession();
  screen = "app";
  render();
  toast(t("hello", { n }));
}
function googleFlow() {
  sheet(`<div class="rowsp" style="margin-bottom:10px"><h2 style="font-size:19px">${t("gTitle")}</h2>
    ${closeBtn()}</div>
    <p class="hint" style="margin:0 0 12px">${t("gSub")}</p>
    <div class="field"><label>${t("email")}</label><input id="g_email" type="email" placeholder="nama@gmail.com"></div>
    <div class="field"><label>${t("fullName")}</label><input id="g_name"></div>
    <button class="btn wide" onclick="doGoogle()">${t("gGo")}</button>`);
}
async function doGoogle() {
  const e = val("g_email").toLowerCase(),
    n = val("g_name") || e.split("@")[0];
  if (!e) return toast(t("fillAll"));
  await pullHub();
  let u = G.users.find((x) => x.email === e);
  if (u) {
    if (!u.active) return toast(t("badLogin"));
    me = u;
    await mutateHub(() => {
      G.users.find((v) => v.email === e).last = now();
    }, "actLogin", "Google");
  } else {
    u = {
      id: uid(),
      name: n,
      email: e,
      role: "treasurer",
      active: true,
      provider: "google",
      pass: "",
      rec: "",
      created: now(),
      last: now(),
      eventIds: [],
    };
    me = u;
    await mutateHub(() => {
      G.users.push(u);
    }, "actSignup", e + " · Google");
  }
  imp = null;
  await enterResolved(me, null);
  closeSheet();
  await saveSession();
  render();
  toast(t("hello", { n: me.name }));
}
async function doForgot() {
  const e = val("r_email").toLowerCase(),
    rc = val("r_rec").toUpperCase(),
    p = val("r_pass");
  if (p.length < 8) return toast(t("passShort"));
  await pullHub();
  const u = G.users.find((x) => x.email === e);
  if (!u || !u.rec || u.rec !== (await sha(e + rc)))
    return toast(t("badRecov"));
  const h = await sha(e + p);
  await mutateHub(() => {
    G.users.find((v) => v.email === e).pass = h;
  }, null);
  toast(t("passReset"));
  authMode = "in";
  render();
}
async function doSignOut(silent) {
  if (!silent) await mutateHub(() => {}, "actLogout", "");
  me = null;
  imp = null;
  currentEventId = null;
  await saveSession();
  screen = "auth";
  authMode = "in";
  render();
}
const eyeOn = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>`;
const eyeOff = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9.9 4.2A10.9 10.9 0 0 1 12 4c6.5 0 10 7 10 7a17 17 0 0 1-3 3.9M6.6 6.6A17 17 0 0 0 2 11s3.5 7 10 7a10.9 10.9 0 0 0 4.1-.8"/><path d="M3 3l18 18"/></svg>`;

function toggleEye(btn) {
  const inp = btn.parentElement.querySelector("input");
  const hidden = inp.type === "password";
  inp.type = hidden ? "text" : "password";
  btn.innerHTML = hidden ? eyeOff : eyeOn;
}

// kode pemulihan dibuat sekali, tetap sama walau layar digambar ulang
let setupRec = "";
function recoveryCode() {
  return Array.from(crypto.getRandomValues(new Uint8Array(4)))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()
    .replace(/(.{4})(.{4})/, "$1-$2"); // contoh: 3F9A-1C0B
}

/* ================= pemilih acara & konsol admin ================= */
function vPicker() {
  const events = myEvents(acting());
  return `<div class="auth"><div class="authbox">
  <div class="rowsp" style="margin-bottom:14px;flex-wrap:wrap;row-gap:10px"><div style="display:flex;gap:10px;align-items:center;min-width:0;flex:1 1 auto">
    <img class="mark" src="/treasurySystem.ico" alt="Treasury System"><div style="min-width:0"><h1 style="font-size:19px">${t("chooseEvent")}</h1>
    <div class="hint">${t("chooseEventSub")}</div></div></div>
    <div style="display:flex;gap:6px;align-items:center;flex:none">${themeBtn()}${langSeg()}</div></div>
  <div class="card" style="padding:16px">
    ${
      events.length
        ? events
            .map(
              (e) => `<button class="btn ghost wide" style="display:flex;flex-direction:column;align-items:flex-start;gap:2px;margin-bottom:8px;padding:14px 16px" onclick="switchEvent('${e.id}')">
        <span style="font-weight:700;font-size:16px">${esc(e.name)}</span>
        ${e.date ? `<span class="hint">${esc(e.date)}</span>` : ""}
      </button>`,
            )
            .join("")
        : `<div class="empty">${t("noEventsAssigned")}</div>`
    }
    <button class="btn danger wide" style="margin-top:8px" onclick="doSignOut()">${t("signOut")}</button>
  </div>${copyrightLine()}</div></div>`;
}
async function switchEvent(id) {
  await enterEvent(id);
  tab = "dash";
  filter = "all";
  txQ = "";
  logQ = "";
  await saveSession();
  closeSheet();
  screen = "app";
  render();
}
function goHub() {
  closeSheet();
  currentEventId = null;
  screen = "hub";
  saveSession();
  render();
}
function setHubTab(k) {
  // keluar dari tab Monitoring membuang cache-nya, supaya lain kali dibuka
  // datanya segar lagi (bukan sisa sebelum ada perubahan acara/staff)
  if (hubTab === "monitor" && k !== "monitor") monitorCache = null;
  hubTab = k;
  render();
}
function vHub() {
  const a = acting();
  return `<header><div class="wrap">
    <img class="mark" src="/treasurySystem.ico" alt="Treasury System">
    <div style="flex:1;min-width:0"><h1 style="font-size:17px">${t("adminConsole")}</h1></div>
    ${themeBtn()}
    ${langSeg()}
    <div class="avatar" onclick="openMe()" title="${esc(a.name)}">${esc(a.name.slice(0, 1).toUpperCase())}</div>
  </div></header>
  <div class="wrap" style="padding:12px 0 50px">
    <div class="seg" style="margin-bottom:12px">${[
      ["events", t("events")],
      ["staff", t("staff")],
      ["monitor", t("monitoring")],
    ]
      .map(
        ([k, l]) =>
          `<button class="${hubTab === k ? "on" : ""}" onclick="setHubTab('${k}')">${l}</button>`,
      )
      .join("")}</div>
    ${
      hubTab === "monitor" && !monitorCache
        ? (loadMonitor(), `<div class="empty">${t("noneYet")}</div>`)
        : { events: vHubEvents, staff: vHubStaff, monitor: vHubMonitor }[hubTab]()
    }
  </div>`;
}
function hubEventsListHtml() {
  const q = hubEventsQ.toLowerCase();
  let all = G.events.filter(
    (e) =>
      (hubEventsFilter === "all" || e.status === hubEventsFilter) &&
      (!q || (e.name + " " + (e.date || "")).toLowerCase().includes(q)),
  );
  const { k, dir } = hubEventsSort,
    mul = dir === "asc" ? 1 : -1;
  all = all.slice().sort((a, b) => {
    const av = String(a[k] ?? "").toLowerCase(),
      bv = String(b[k] ?? "").toLowerCase();
    return av > bv ? mul : av < bv ? -mul : 0;
  });
  const { items, page, totalPages } = paginate(all, hubEventsPage, 10);
  if (!items.length) return `<div class="empty">${t("noneYet")}</div>`;
  if (narrow())
    return (
      items
        .map(
          (e) => `<div class="drill-row">
      <div class="drill-row-bar" style="background:${e.status === "active" ? "var(--green)" : "var(--red)"}"></div>
      <div style="flex:1;min-width:0">
        <div class="rowsp" style="gap:8px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(e.name)}</div>
            <div class="hint mono" style="margin-top:2px">${e.date || "-"}</div>
          </div>
          <span class="tag ${e.status === "active" ? "t-ok" : "t-exp"}" style="flex:none">${e.status === "active" ? t("active") : t("archived")}</span>
        </div>
        <div class="rowsp" style="margin-top:8px;gap:8px;flex-wrap:wrap">
          ${e.status === "active" ? `<button class="btn ghost sm" onclick="switchEvent('${e.id}')">${t("enterWorkspace")}</button>` : "<span></span>"}
          <div style="display:flex;gap:8px">
            <button class="btn ghost sm" onclick="openDuplicateEvent('${e.id}')">${t("duplicate")}</button>
            <button class="btn ghost sm" onclick="toggleArchive('${e.id}')">${e.status === "active" ? t("archive") : t("restore")}</button>
          </div>
        </div>
      </div></div>`,
        )
        .join("") + pagerHtml(page, totalPages, "setHubEventsPage")
    );
  const si = (k) =>
    hubEventsSort.k === k ? (hubEventsSort.dir === "asc" ? " ▲" : " ▼") : "";
  return `<div style="overflow-x:auto"><table><thead><tr>
      <th class="sortable" onclick="setHubEventsSort('name')">${t("name")}${si("name")}</th>
      <th class="sortable" onclick="setHubEventsSort('date')">${t("evDate")}${si("date")}</th>
      <th class="sortable" onclick="setHubEventsSort('status')">${t("status")}${si("status")}</th>
      <th></th></tr></thead><tbody>
    ${items
      .map(
        (e) => `<tr>
      <td style="font-weight:600">${esc(e.name)}</td>
      <td class="hint mono">${e.date || "-"}</td>
      <td><span class="tag ${e.status === "active" ? "t-ok" : "t-exp"}">${e.status === "active" ? t("active") : t("archived")}</span></td>
      <td style="text-align:right;white-space:nowrap">
        ${e.status === "active" ? `<button class="btn ghost sm" onclick="switchEvent('${e.id}')">${t("enterWorkspace")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openDuplicateEvent('${e.id}')">${t("duplicate")}</button>
        <button class="btn ghost sm" onclick="toggleArchive('${e.id}')">${e.status === "active" ? t("archive") : t("restore")}</button>
      </td></tr>`,
      )
      .join("")}
    </tbody></table></div>${pagerHtml(page, totalPages, "setHubEventsPage")}`;
}
function vHubEvents() {
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("events")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;flex-wrap:wrap;min-width:220px">
      <div class="chips">${[
        ["all", t("all")],
        ["active", t("active")],
        ["archived", t("archived")],
      ]
        .map(
          ([k2, l]) =>
            `<button class="chip ${hubEventsFilter === k2 ? "on" : ""}" onclick="setHubEventsFilter('${k2}')">${l}</button>`,
        )
        .join("")}</div>
      <input id="hubEventsSearchInput" style="max-width:220px" placeholder="${t("searchTx")}" value="${esc(hubEventsQ)}" oninput="onHubEventsSearch(this.value)">
      <button class="btn sm" onclick="openNewEvent()">${t("newEvent")}</button></div></div>
    <div id="hubEventsListBody">${hubEventsListHtml()}</div></div>`;
}
function setHubEventsSort(k) {
  hubEventsSort =
    hubEventsSort.k === k
      ? { k, dir: hubEventsSort.dir === "asc" ? "desc" : "asc" }
      : { k, dir: "asc" };
  hubEventsPage = 1;
  render();
}
function openNewEvent() {
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${t("newEvent")}</h2>${closeBtn()}</div>
    <div class="field"><label>${t("evName")}</label><input id="ne_name"></div>
    <div class="field"><label>${t("evDate")}</label><input type="date" id="ne_date"></div>
    <button class="btn wide" onclick="createEventSubmit()">${t("saveGeneric")}</button>`);
}
async function createEventSubmit() {
  const name = val("ne_name"),
    date = val("ne_date");
  if (!name) return toast(t("fillAll"));
  const entry = await createEventRow("ev_" + uid(), name, date, acting().email);
  await mutateHub(() => {
    G.events.push(entry);
  }, "actEventCreate", name);
  closeSheet();
  toast(t("saved"));
  render();
}
// beda dari createEventRow() (yg selalu mulai kosong) - baris baru ini
// mewarisi cats/methods/tpl/dashboard PERSIS dari acara sumber, tapi tx &
// logs selalu mulai kosong (bukan acara sungguhan, jadi tidak boleh bawa
// riwayat transaksi acara lain)
async function duplicateEventRow(srcId, newId, name, date, createdBy) {
  const r = await window.storage.get(eventKey(srcId), true);
  const src = r && r.value ? JSON.parse(r.value) : null;
  const config = src
    ? { ...src.config, event: name, date: date || "" }
    : { event: name, date: date || "", cats: [], methods: [], tpl: DEFAULT_TPL, dashboard: { widgets: starterDashboardWidgets() } };
  await window.storage.set(eventKey(newId), JSON.stringify({ rev: 1, config, tx: [], logs: [] }), true);
  return { id: newId, name, date: date || "", status: "active", createdAt: now(), createdBy };
}
function openDuplicateEvent(id) {
  const src = G.events.find((e) => e.id === id);
  if (!src) return;
  const suggested = src.name + (lang === "id" ? " (Salinan)" : " (Copy)");
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${t("duplicateEvent")}</h2>${closeBtn()}</div>
    <p class="hint" style="margin:0 0 12px">${t("duplicateEventHint", { n: esc(src.name) })}</p>
    <input type="hidden" id="dup_src" value="${id}">
    <div class="field"><label>${t("evName")}</label><input id="dup_name" value="${esc(suggested)}"></div>
    <div class="field"><label>${t("evDate")}</label><input type="date" id="dup_date" value="${src.date || ""}"></div>
    <button class="btn wide" onclick="submitDuplicateEvent()">${t("duplicateEvent")}</button>`);
}
async function submitDuplicateEvent() {
  const srcId = val("dup_src"),
    name = val("dup_name"),
    date = val("dup_date");
  if (!name) return toast(t("fillAll"));
  const srcName = G.events.find((e) => e.id === srcId)?.name || "";
  const entry = await duplicateEventRow(srcId, "ev_" + uid(), name, date, acting().email);
  await mutateHub(() => {
    G.events.push(entry);
  }, "actEventDuplicate", `${srcName} -> ${name}`);
  closeSheet();
  toast(t("saved"));
  render();
}
async function toggleArchive(id) {
  const e = G.events.find((x) => x.id === id);
  const next = e.status === "active" ? "archived" : "active";
  await mutateHub(() => {
    G.events.find((x) => x.id === id).status = next;
  }, "actEventArchive", e.name + " -> " + next);
  render();
}
// mengambil data tiap acara sekaligus, hanya saat tab Monitoring dibuka -
// bukan polling, supaya tidak membaca semua acara tiap 6 detik
async function loadMonitor() {
  const rows = await Promise.all(
    G.events.map(async (e) => {
      const r = await window.storage.get(eventKey(e.id), true);
      const data =
        r && r.value ? JSON.parse(r.value) : { config: { cats: [] }, tx: [], logs: [] };
      return { event: e, data };
    }),
  );
  monitorCache = {
    perEvent: rows.map(({ event, data }) => ({
      event,
      sums: sums(data.tx, data.config.cats),
    })),
    pending: rows
      .flatMap(({ event, data }) =>
        data.tx
          .filter((x) => x.status === "pending")
          .map((x) => ({ ...x, eventId: event.id, eventName: event.name })),
      )
      .sort((a, b) => a.date.localeCompare(b.date)),
    staff: G.users
      .filter((u) => u.role === "treasurer")
      .map((u) => {
        const mapped = rows.filter(({ event }) => (u.eventIds || []).includes(event.id));
        return {
          user: u,
          recorded: mapped.reduce(
            (n, { data }) => n + data.tx.filter((x) => x.by === u.name).length,
            0,
          ),
          verified: mapped.reduce(
            (n, { data }) => n + data.tx.filter((x) => x.vBy === u.name).length,
            0,
          ),
          events: mapped.map(({ event }) => event.name),
        };
      }),
  };
  render();
}
function vHubMonitor() {
  const m = monitorCache;
  if (!m) return `<div class="empty">${t("noneYet")}</div>`;
  const k = (l, v, c) =>
    `<div class="card kpi act" style="padding:11px 13px;justify-content:center"><div class="label">${l}</div><div class="n mono ${c || ""}">${v}</div></div>`;
  return `<div class="rowsp" style="margin-bottom:10px">
    <h2 style="font-size:17px">${t("financialSummary")}</h2>
    <button class="btn ghost sm" onclick="loadMonitor()">${t("refresh")}</button></div>
  <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-bottom:16px">
    ${
      m.perEvent
        .map(
          ({ event, sums: s }) => `<div class="card" style="padding:13px">
      <div class="rowsp" style="margin-bottom:6px"><span style="font-weight:700">${esc(event.name)}</span>
        <span class="tag ${event.status === "active" ? "t-ok" : "t-exp"}">${event.status === "active" ? t("active") : t("archived")}</span></div>
      <div class="grid" style="grid-template-columns:1fr 1fr;gap:6px">
        ${k(t("net"), rpk(s.net), s.net >= 0 ? "pos" : "neg")}
        ${k(t("wait"), rpk(s.pendAmt), "amb")}
      </div></div>`,
        )
        .join("") || `<div class="empty">${t("noneYet")}</div>`
    }
  </div>
  <div class="card" style="margin-bottom:16px"><h2 style="font-size:17px;margin-bottom:8px">${t("needsAttention")}</h2>
    ${
      m.pending.length
        ? `<div class="scroll" style="max-height:320px">${m.pending
            .map(
              (x) => `<div class="rowsp" style="padding:7px 0;border-top:1px solid var(--line2)">
        <div style="min-width:0"><div style="font-weight:600">${esc(x.name)}</div>
        <div class="hint">${esc(x.eventName)} · ${rpk(x.amount)}</div></div>
        <button class="btn ghost sm" onclick="switchEvent('${x.eventId}')">${t("enterWorkspace")}</button></div>`,
            )
            .join("")}</div>`
        : `<div class="empty">${t("noQueue")}</div>`
    }</div>
  <div class="card"><h2 style="font-size:17px;margin-bottom:8px">${t("staffPerformance")}</h2>
    <div style="overflow-x:auto"><table style="min-width:520px"><thead><tr><th style="min-width:140px">${t("name")}</th><th style="min-width:180px">${t("events")}</th><th style="min-width:90px">${t("recorded")}</th><th style="min-width:90px">${t("verified")}</th></tr></thead><tbody>
    ${
      m.staff
        .map(
          (s) => `<tr><td style="font-weight:600">${esc(s.user.name)}</td>
      <td class="hint">${s.events.map(esc).join(", ") || "-"}</td>
      <td class="mono">${s.recorded}</td>
      <td class="mono">${s.verified}</td></tr>`,
        )
        .join("") || `<tr><td colspan="4" class="empty">${t("noneYet")}</td></tr>`
    }
    </tbody></table></div></div>`;
}

/* ================= kerangka aplikasi ================= */
function vApp() {
  const s = sums(),
    a = acting();
  const hasQtyCats = D().cats.some((c) => c.group === "income" && c.hasQty);
  const headerSub = [
    D().date
      ? new Date(D().date + "T00:00:00").toLocaleDateString(lang === "id" ? "id-ID" : "en-GB", {
          day: "numeric",
          month: "short",
          year: "numeric",
        })
      : "",
    hasQtyCats ? s.sold + " " + t("seatsSold") : "",
  ]
    .filter(Boolean)
    .join(" · ");
  // viewer: cuma boleh lihat Dashboard - tidak ada Transaksi/Peringkat/Admin
  // di nav, spy tidak ada jalan masuk ke layar yg bisa mengubah data
  const tabs = canEdit() ? [["dash", t("dash")], ["tx", t("tx")], ["board", t("board")]] : [["dash", t("dash")]];
  if (isAdmin()) tabs.push(["admin", t("admin")]);
  if (tab === "admin" && !isAdmin()) tab = "dash";
  if ((tab === "tx" || tab === "board") && !canEdit()) tab = "dash";
  return `${
    imp
      ? `<div class="banner">${t("impBanner", { n: esc(imp.name) })}
      <button class="btn sm" onclick="stopImp()">${t("backAdmin")}</button></div>`
      : ""
  }
  <header><div class="wrap">
    <img class="mark" src="/treasurySystem.ico" alt="Treasury System">
    <div style="flex:1;min-width:0"><h1 style="font-size:17px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(D().event || "Konser")}</h1>
      <div class="sub">${esc(headerSub)}</div></div>
    ${themeBtn()}
    ${langSeg()}
    <button class="btn ghost sm xls-btn" onclick="openExportPeriod()">${downloadIcon}<span class="xls-label">${t("xls")}</span></button>
    <div class="avatar" onclick="openMe()" title="${esc(a.name)}">${esc(a.name.slice(0, 1).toUpperCase())}</div>
  </div></header>
  <div class="wrap">${{ dash: vDash, tx: vTx, board: vBoard, admin: vAdmin }[tab]()}</div>
  ${(tab === "tx" || tab === "dash") && canEdit() ? `<button class="btn fab" onclick="openTx()">${t("add")}</button>` : ""}
  <nav>${tabs.map(([k, l]) => `<button class="${tab === k ? "on" : ""}" onclick="go('${k}')"><span class="dot"></span>${l}</button>`).join("")}</nav>`;
}
function go(k) {
  tab = k;
  render();
}
function goDashEditor() {
  tab = "admin";
  adminTab = "dashboard";
  render();
}
function setAuthMode(f) {
  authMode = f;
  render();
}
function setPanel(p) {
  panel = p;
  render();
}
function setFilter(k) {
  filter = k;
  txPage = 1;
  render();
}
function setAdminTab(k) {
  adminTab = k;
  render();
}
function onLogSearch(v) {
  logQ = v;
  logPage = 1;
  clearTimeout(window._lq);
  window._lq = setTimeout(() => {
    const el = document.getElementById("logListBody");
    if (el) el.innerHTML = logListHtml();
  }, 250);
}
function setLogPage(p) {
  logPage = p;
  render();
}
function setTxPage(p) {
  txPage = p;
  render();
}
function setBoardBuyPage(p) {
  boardBuyPage = p;
  render();
}
function setBoardDonPage(p) {
  boardDonPage = p;
  render();
}
function setHubEventsPage(p) {
  hubEventsPage = p;
  render();
}
function setHubEventsFilter(k) {
  hubEventsFilter = k;
  hubEventsPage = 1;
  render();
}
function onHubEventsSearch(v) {
  hubEventsQ = v;
  hubEventsPage = 1;
  clearTimeout(window._heq);
  window._heq = setTimeout(() => {
    const el = document.getElementById("hubEventsListBody");
    if (el) el.innerHTML = hubEventsListHtml();
  }, 250);
}
function setHubStaffPage(p) {
  hubStaffPage = p;
  render();
}
function onHubStaffSearch(v) {
  hubStaffQ = v;
  hubStaffPage = 1;
  clearTimeout(window._hsq);
  window._hsq = setTimeout(() => {
    const el = document.getElementById("hubStaffListBody");
    if (el) el.innerHTML = hubStaffListHtml();
  }, 250);
}
function setHubStaffSort(k) {
  hubStaffSort =
    hubStaffSort.k === k
      ? { k, dir: hubStaffSort.dir === "asc" ? "desc" : "asc" }
      : { k, dir: "asc" };
  hubStaffPage = 1;
  render();
}
function openMe() {
  const a = acting();
  const events = myEvents(a).filter((e) => e.id !== currentEventId);
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${esc(a.name)}</h2>
    ${closeBtn()}</div>
    <div class="hint" style="margin-bottom:6px">${esc(a.email)}</div>
    <span class="tag ${roleTagClass(a.role)}">${roleLabel(a.role)}</span>
    ${
      imp
        ? `<p class="hint" style="margin-top:12px">${t("byAdmin", { n: esc(me.name) })}</p>
      <button class="btn wide" style="margin-top:8px" onclick="stopImp()">${t("backAdmin")}</button>`
        : ""
    }
    ${
      events.length
        ? `<div class="field" style="margin-top:14px">
      <label>${t("switchEvent")}</label>
      <select id="f_switchEvent" onchange="this.value && switchEvent(this.value)">
        <option value="" selected disabled>${t("chooseEvent")}</option>
        ${events.map((e) => `<option value="${e.id}">${esc(e.name)}</option>`).join("")}
      </select></div>`
        : ""
    }
    ${
      isAdmin() && screen === "app"
        ? `<button class="btn ghost wide" style="margin-top:6px" onclick="goHub()">${t("adminConsole")}</button>`
        : ""
    }
    <button class="btn danger wide" style="margin-top:14px" onclick="closeSheet();doSignOut()">${t("signOut")}</button>`);
  if (events.length) initCombobox("f_switchEvent", t("chooseEvent"));
}
async function stopImp() {
  const n = imp.name;
  imp = null;
  closeSheet();
  await saveSession();
  render();
  toast(t("backAdmin") + " · " + n);
}

/* ================= dashboard (widget) ================= */
// grid dashboard: tiap widget punya posisi/ukuran (x,y,w,h) dalam satuan
// grid CSS 12-kolom. Mode lihat (di sini) murni CSS grid, tanpa lib apa pun -
// editor drag/resize-nya (gridstack, admin-only) baru dimuat lazy di tab
// terpisah, lihat vDashEditor()/goDashEditor().
const DASHBOARD_COLS = 12;
const DASHBOARD_ROW_H = 80;
const PIE_COLORS = [
  "--chart-1", "--chart-2", "--chart-3", "--chart-4",
  "--chart-5", "--chart-6", "--chart-7", "--chart-8",
];
// tampilkan label persentase langsung di tiap slice (bukan cuma pas hover) -
// posisinya diambil dari getCenterPoint() bawaan ArcElement Chart.js, jadi
// otomatis benar utk pie maupun doughnut, apa pun ukuran cutout-nya
const pieLabelsPlugin = {
  id: "pieLabels",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;
      const total = dataset.data.reduce((a, b) => a + b, 0);
      if (!total) return;
      meta.data.forEach((arc, i) => {
        const v = dataset.data[i];
        const pct = Math.round((v / total) * 100);
        if (!v || pct < 4) return; // slice terlalu tipis, label akan tumpang tindih
        const pos = arc.getCenterPoint();
        ctx.save();
        ctx.fillStyle = "#fff";
        ctx.font = "700 12px Inter, system-ui, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.shadowColor = "rgba(0,0,0,.45)";
        ctx.shadowBlur = 4;
        ctx.fillText(pct + "%", pos.x, pos.y);
        ctx.restore();
      });
    });
  },
};
// judul widget default (dibuat via t() saat migrasi/starter/tambah widget)
// disimpan sbg teks statis di config.dashboard.widgets - kalau bahasa
// diganti belakangan, teks yg SUDAH TERSIMPAN itu tidak otomatis berubah.
// Supaya konsisten, kalau judul yg tersimpan masih persis sama dgn salah
// satu versi bahasa aslinya (belum diedit manual oleh admin jadi judul
// custom), tampilkan terjemahan LIVE-nya saja, bukan teks beku itu.
const WIDGET_TITLE_KEYS = [
  "net", "incomeW", "exp", "wait", "seats", "queue", "topBuy", "topDon", "rank",
  "wKpi", "wChart", "wPie", "wTable", "wBreakdown", "wQuota", "wQueue", "wRank",
];
function liveTitle(storedTitle) {
  for (const key of WIDGET_TITLE_KEYS) {
    if (storedTitle === L.id[key] || storedTitle === L.en[key]) return t(key);
  }
  return storedTitle;
}
const widgetTitle = (w) => `<span class="label">${esc(liveTitle(w.title))}</span>`;
function widgetCatsInGroup(w) {
  // group tak diisi (dipakai widget "table" generik) -> semua kategori
  const all = w.group ? D().cats.filter((c) => c.group === w.group) : D().cats;
  return w.catIds && w.catIds.length ? all.filter((c) => w.catIds.includes(c.id)) : all;
}
function widgetKpiValue(w, s) {
  if (w.metric === "net") return s.net;
  if (w.metric === "income") return s.income;
  if (w.metric === "expense") return s.exp;
  if (w.metric === "pendingAmt") return s.pendAmt;
  if (typeof w.metric === "string" && w.metric.startsWith("cat:")) {
    const c = D().cats.find((x) => x.id === w.metric.slice(4));
    return c ? s.byCat[c.n] || 0 : 0;
  }
  return 0;
}

/* ----- drill-down: klik kartu/baris/titik grafik -> pop up daftar transaksinya ----- */
const DRILL_PAGE_SIZE = 8;
let drillState = null;
function drillMatches(x, q) {
  if (!q) return true;
  const hay = [x.name, catLabel(x), rp(x.amount), String(x.amount), x.status === "verified" ? t("inRek") : t("stW"), x.note]
    .join(" ")
    .toLowerCase();
  return q
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .every((term) => hay.includes(term));
}
function drillRowHtml(x) {
  const isExp = x.type === "expense";
  return `<div class="drill-row drill" onclick="closeSheet();openTx('${x.id}')">
    <div class="drill-row-bar ${isExp ? "drill-bar-neg" : "drill-bar-pos"}"></div>
    <div style="flex:1;min-width:0;display:flex;justify-content:space-between;gap:10px;align-items:flex-start">
      <div style="min-width:0">
        <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(x.name)}</div>
        <div class="hint" style="margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(catLabel(x))} · ${x.date.slice(8)}/${x.date.slice(5, 7)}/${x.date.slice(0, 4)}</div>
      </div>
      <div style="text-align:right;flex:none">
        <div class="mono" style="font-weight:700${isExp ? ";color:var(--red)" : ""}">${isExp ? "−" : ""}${rp(x.amount)}</div>
        <span class="tag ${x.status === "verified" ? "t-ok" : "t-wait"}" style="margin-top:3px;display:inline-block">${x.status === "verified" ? t("inRek") : t("stW")}</span>
      </div>
    </div>
  </div>`;
}
function renderDrillBody() {
  const filtered = drillState.txList
    .filter((x) => drillMatches(x, drillState.q))
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date));
  const { items, page, totalPages } = paginate(filtered, drillState.page, DRILL_PAGE_SIZE);
  drillState.page = page;
  const total = filtered.reduce((a, x) => a + (x.type === "expense" ? -x.amount : x.amount), 0);
  const statsEl = document.getElementById("drillStats");
  if (statsEl)
    statsEl.innerHTML = `
    <div class="drill-stat"><b>${filtered.length}</b><span class="hint">${t("txCountLbl")}</span></div>
    <div class="drill-stat"><b class="mono">${rp(total)}</b><span class="hint">${t("totalLbl")}</span></div>`;
  const bodyEl = document.getElementById("drillBody");
  if (bodyEl)
    bodyEl.innerHTML = items.length
      ? items.map(drillRowHtml).join("") + pagerHtml(page, totalPages, "setDrillPage")
      : `<div class="empty">${t("noneYet")}</div>`;
}
function onDrillSearch(v) {
  drillState.q = v;
  drillState.page = 1;
  renderDrillBody();
}
function setDrillPage(p) {
  drillState.page = p;
  renderDrillBody();
}
function openTxListSheet(title, txList) {
  drillState = { title, txList, q: "", page: 1 };
  sheet(`<div class="rowsp" style="margin-bottom:10px"><h2 style="font-size:19px">${esc(title)}</h2>${closeBtn()}</div>
    <div class="drill-stats" id="drillStats"></div>
    <input id="drillSearch" placeholder="${t("drillSearchPlaceholder")}" style="margin-bottom:12px" oninput="onDrillSearch(this.value)">
    <div id="drillBody" style="max-height:56vh;overflow-y:auto;padding-bottom:8px"></div>`);
  renderDrillBody();
}
// "Saldo bersih" tidak masuk akal ditampilkan sbg daftar transaksi datar -
// yg berguna adalah saldo KUMULATIF per hari (H vs H-1) beserta persentase
// naik/turunnya, jadi widget ini dapat tampilan drill-down sendiri
function balanceAsOf(dateStr) {
  return S.tx
    .filter((x) => x.status === "verified" && x.date <= dateStr)
    .reduce((a, x) => a + (x.type === "expense" ? -x.amount : x.amount), 0);
}
function drilldownNet() {
  const N = 14;
  const days = [];
  for (let i = N - 1; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    days.push(d.toISOString().slice(0, 10));
  }
  const bal = days.map(balanceAsOf);
  const rows = days
    .map((d, i) => {
      const prev = i > 0 ? bal[i - 1] : null;
      const delta = prev !== null ? bal[i] - prev : null;
      const pct = prev ? (delta / Math.abs(prev)) * 100 : delta && delta !== 0 ? 100 : 0;
      return { date: d, balance: bal[i], delta, pct };
    })
    .reverse();
  sheet(`<div class="rowsp" style="margin-bottom:4px"><h2 style="font-size:19px">${t("net")}</h2>${closeBtn()}</div>
    <p class="hint" style="margin:0 0 12px">${t("dailyBalanceHint")}</p>
    <div style="overflow-x:auto;max-height:65vh;overflow-y:auto"><table><thead><tr>
      <th>${t("date")}</th><th style="text-align:right">${t("net")}</th><th style="text-align:right">${t("changeLbl")}</th></tr></thead><tbody>
    ${rows
      .map((r) => {
        const flat = r.delta === 0;
        const up = r.delta === null || flat ? null : r.delta >= 0;
        const arrow = up === null ? "" : up ? trendUpIcon : trendDownIcon;
        const cls = up === null ? "hint" : up ? "pos" : "neg";
        const dt = new Date(r.date + "T00:00:00").toLocaleDateString(lang === "id" ? "id-ID" : "en-GB", {
          day: "2-digit",
          month: "short",
        });
        const changeText = r.delta === null ? "—" : flat ? "0.0%" : `${r.pct >= 0 ? "+" : ""}${r.pct.toFixed(1)}%`;
        return `<tr><td class="mono hint">${dt}</td>
        <td class="mono" style="text-align:right;font-weight:600">${rp(r.balance)}</td>
        <td class="${cls}" style="text-align:right;white-space:nowrap"><span style="display:inline-flex;align-items:center;gap:4px;justify-content:flex-end">${arrow}${changeText}</span></td></tr>`;
      })
      .join("")}
    </tbody></table></div>`);
}
function txForCat(catName, verifiedOnly = true) {
  return S.tx.filter((x) => x.cat === catName && (!verifiedOnly || x.status === "verified"));
}
function txForPerson(name, catNames) {
  return S.tx.filter(
    (x) =>
      x.status === "verified" &&
      x.type === "income" &&
      x.name === name &&
      (!catNames || !catNames.length || catNames.includes(x.cat)),
  );
}
function txForDate(dateStr) {
  return S.tx.filter((x) => x.date === dateStr && x.status === "verified" && x.type === "income");
}
function txForKpi(w) {
  if (w.metric === "income") return S.tx.filter((x) => x.status === "verified" && x.type === "income");
  if (w.metric === "expense") return S.tx.filter((x) => x.status === "verified" && x.type === "expense");
  if (w.metric === "net") return S.tx.filter((x) => x.status === "verified");
  if (w.metric === "pendingAmt") return S.tx.filter((x) => x.status === "pending");
  if (typeof w.metric === "string" && w.metric.startsWith("cat:")) {
    const c = D().cats.find((x) => x.id === w.metric.slice(4));
    return c ? txForCat(c.n) : [];
  }
  return [];
}
function findWidget(id) {
  return (S.config.dashboard?.widgets || []).find((w) => w.id === id);
}
function drilldownKpi(widgetId) {
  const w = findWidget(widgetId);
  if (!w) return;
  if (w.metric === "net") return drilldownNet();
  openTxListSheet(liveTitle(w.title), txForKpi(w));
}
function drilldownCat(catName) {
  openTxListSheet(catName, txForCat(catName, false));
}
function drilldownPerson(name, widgetId) {
  const w = findWidget(widgetId);
  openTxListSheet(name, txForPerson(name, w ? widgetCatsInGroup(w).map((c) => c.n) : null));
}
function drilldownDate(dateStr) {
  openTxListSheet(
    new Date(dateStr + "T00:00:00").toLocaleDateString(lang === "id" ? "id-ID" : "en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric",
    }),
    txForDate(dateStr),
  );
}

// bandingkan nilai KUMULATIF per metrik hari ini vs kemarin (H vs H-1) -
// "pendingAmt" dilewati krn belum lunas bukan konsep yg akumulatif dari waktu
function widgetKpiTrend(w) {
  if (w.metric === "pendingAmt") return null;
  const yesterday = (() => {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    return d.toISOString().slice(0, 10);
  })();
  const valueAsOf = (dateStr) => {
    const tx = S.tx.filter((x) => x.status === "verified" && x.date <= dateStr);
    if (w.metric === "income") return tx.filter((x) => x.type === "income").reduce((a, b) => a + b.amount, 0);
    if (w.metric === "expense") return tx.filter((x) => x.type === "expense").reduce((a, b) => a + b.amount, 0);
    if (w.metric === "net") return tx.reduce((a, x) => a + (x.type === "expense" ? -x.amount : x.amount), 0);
    if (typeof w.metric === "string" && w.metric.startsWith("cat:")) {
      const c = D().cats.find((x) => x.id === w.metric.slice(4));
      return c ? tx.filter((x) => x.cat === c.n).reduce((a, b) => a + b.amount, 0) : null;
    }
    return null;
  };
  const vToday = valueAsOf(today()),
    vYesterday = valueAsOf(yesterday);
  if (vToday === null || vYesterday === null) return null;
  const delta = vToday - vYesterday;
  if (!delta) return { delta: 0, pct: 0, up: true, flat: true };
  const pct = vYesterday !== 0 ? (delta / Math.abs(vYesterday)) * 100 : 100;
  return { delta, pct, up: delta >= 0, flat: false };
}
function renderKpiWidget(w, s) {
  const val = widgetKpiValue(w, s);
  const cls =
    w.metric === "expense" ? "neg" : w.metric === "net" ? (val >= 0 ? "pos" : "neg") : "blu";
  const trend = widgetKpiTrend(w);
  // warna badge tren selalu ikut arah matematisnya (naik=hijau, turun=merah)
  // - tidak dibalik utk "pengeluaran naik = buruk", biar konsisten & tidak
  // membingungkan (naik selalu hijau, apa pun metriknya)
  const trendHtml = trend
    ? `<div class="kpi-trend ${trend.flat ? "flat" : trend.up ? "pos" : "neg"}" title="${t("vsYesterday")}">${trend.flat ? "" : trend.up ? trendUpIcon : trendDownIcon}<span>${trend.flat ? "0%" : `${trend.pct >= 0 ? "+" : ""}${trend.pct.toFixed(1)}%`}</span></div>`
    : "";
  return `<div class="rowsp" style="align-items:flex-start;flex:none;flex-wrap:wrap;row-gap:4px"><div style="min-width:0"><div class="label">${esc(liveTitle(w.title))}</div><div class="n mono ${cls}">${rpk(val)}</div></div>${trendHtml}</div>`;
}
function renderChartWidget(w) {
  return `<div class="rowsp" style="flex:none">${widgetTitle(w)}
    <span class="mono" style="font-weight:700;font-size:17px">${rpk(sums().income)}</span></div>
    <div class="chartbox"><canvas id="wchart_${w.id}" data-chart-widget="${w.id}"></canvas></div>`;
}
function renderPieWidget(w) {
  return `${widgetTitle(w)}<div class="chartbox"><canvas id="wpie_${w.id}" data-pie-widget="${w.id}"></canvas></div>`;
}
function renderTableWidget(w) {
  const catNames = widgetCatsInGroup(w).map((c) => c.n);
  const all = S.tx
    .filter((x) => catNames.includes(x.cat))
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date));
  const rows = all.slice(0, 30);
  return `${widgetTitle(w)}<div class="scroll" style="flex:1;margin-top:4px"><table style="font-size:13px"><thead><tr>
    <th>${t("date")}</th><th>${t("name")}</th><th style="text-align:right">${t("amount")}</th></tr></thead><tbody>
    ${
      rows
        .map(
          (x) => `<tr class="drill" onclick="openTx('${x.id}')" title="${esc(catLabel(x))} · ${x.status === "verified" ? t("inRek") : t("stW")}">
      <td class="mono hint" style="white-space:nowrap">${x.date.slice(8)}/${x.date.slice(5, 7)}</td>
      <td style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px">${esc(x.name)}</td>
      <td class="mono ${x.type === "expense" ? "neg" : ""}" style="text-align:right;white-space:nowrap">${x.type === "expense" ? "−" : ""}${rp(x.amount)}</td></tr>`,
        )
        .join("") || `<tr><td colspan="3" class="empty">${t("noneYet")}</td></tr>`
    }
    </tbody></table>${all.length > 30 ? `<div class="hint" style="text-align:center;padding-top:6px">${t("tableMoreHint", { n: all.length - 30 })}</div>` : ""}</div>`;
}
function renderBreakdownWidget(w, s) {
  const cats = widgetCatsInGroup(w)
    .slice()
    .sort((a, b) => (s.byCat[b.n] || 0) - (s.byCat[a.n] || 0));
  const total = cats.reduce((a, c) => a + (s.byCat[c.n] || 0), 0);
  return `${widgetTitle(w)}<div class="scroll" style="flex:1;margin-top:4px">${
    cats
      .map((c) => {
        const v = s.byCat[c.n] || 0,
          pct = total ? Math.round((v / total) * 100) : 0;
        return `<div class="rowsp drill" style="padding:6px 0;border-top:1px solid var(--line2);font-size:15px" onclick="drilldownCat('${esc(c.n)}')" title="${pct}% ${t("ofTotal")} · ${rp(v)}">
      <span>${esc(c.n)}</span><span class="mono" style="font-weight:600">${rp(v)}</span></div>`;
      })
      .join("") || `<div class="hint" style="padding-top:10px">${t("noneYet")}</div>`
  }</div>`;
}
function renderQuotaWidget(w) {
  const cats = D().cats.filter(
    (c) => c.group === "income" && c.hasQty && (!w.catIds?.length || w.catIds.includes(c.id)),
  );
  const quota = cats.reduce((a, c) => a + (+c.q || 0), 0);
  const soldOf = (c) =>
    S.tx
      .filter((x) => x.status === "verified" && x.cat === c.n)
      .reduce((a, b) => a + (+b.seats || 0), 0);
  const sold = cats.reduce((a, c) => a + soldOf(c), 0);
  const pct = quota ? Math.round((sold / quota) * 100) : 0;
  return `${widgetTitle(w)}
  <div class="mono" style="font-size:26px;font-weight:700;letter-spacing:-.03em;margin-top:4px">${Math.max(quota - sold, 0)}<span style="font-size:13px;color:var(--tx2);font-weight:500"> ${t("left")}</span></div>
  <div class="bar"><i style="width:${Math.min(pct, 100)}%"></i></div>
  <div class="rowsp hint" style="flex:none"><span>${t("soldOf", { a: sold, b: quota })}</span><span>${pct}%</span></div>
  <div class="scroll" style="margin-top:6px;flex:1">${
    cats
      .map((c) => {
        const sd = soldOf(c),
          v = rp((D().cats.find((x) => x.n === c.n) || {}).p * sd || 0);
        return `<div class="rowsp drill" style="padding:6px 0;border-top:1px solid var(--line2);font-size:14px" onclick="drilldownCat('${esc(c.n)}')" title="${t("txCat")}: ${esc(c.n)} · ${v}">
      <span>${esc(c.n)}</span><span class="mono" style="font-weight:600">${sd}/${c.q || "∞"}</span></div>`;
      })
      .join("") || `<div class="hint" style="padding-top:10px">${t("noneYet")}</div>`
  }</div>`;
}
function renderQueueWidget(w) {
  const q = S.tx.filter((x) => x.status === "pending");
  return `${widgetTitle(w)}<div class="scroll" style="flex:1;margin-top:4px">${
    q
      .map(
        (x) => `<div class="rowsp drill" style="padding:7px 0;border-top:1px solid var(--line2)" onclick="openTx('${x.id}')">
    <div style="min-width:0"><div style="font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(x.name)}</div>
    <div class="hint mono">${rpk(x.amount)} · ${esc(x.bank || "-")}</div></div>
    ${canEdit() ? `<button class="btn ok sm" onclick="event.stopPropagation();verify('${x.id}')">${t("inRek")}</button>` : `<span class="tag t-wait">${t("stW")}</span>`}</div>`,
      )
      .join("") || `<div class="hint" style="padding-top:10px">${t("noQueue")}</div>`
  }</div>`;
}
const MEDAL_COLORS = ["#e8b923", "#a9adb4", "#c4813f"];
function renderRankWidget(w) {
  const names = widgetCatsInGroup(w).map((c) => c.n);
  const rows = byPerson(names).slice(0, 8);
  const max = rows.reduce((m, x) => Math.max(m, x.amount), 0) || 1;
  return `${widgetTitle(w)}<div class="scroll" style="flex:1;margin-top:8px">${
    rows
      .map((x, i) => {
        const pct = Math.max(6, Math.round((x.amount / max) * 100));
        const medal = i < 3 ? MEDAL_COLORS[i] : null;
        const detail = `${x.n} ${t("trans")}${x.seats ? " · " + x.seats + " " + t("seatsW").toLowerCase() : ""}`;
        return `<div class="rank-row drill" onclick="drilldownPerson('${esc(x.name).replace(/'/g, "&#39;")}','${w.id}')" title="${detail} · ${rp(x.amount)}">
      <div class="rank-badge" style="${medal ? `background:${medal};color:#fff` : "background:var(--line2);color:var(--tx2)"}">${i + 1}</div>
      <div style="flex:1;min-width:0">
        <div class="rowsp" style="gap:8px"><span class="rank-name">${esc(x.name)}</span><span class="mono rank-amt">${rpk(x.amount)}</span></div>
        <div class="rank-bar"><i style="width:${pct}%;background:${medal || "var(--blue)"}"></i></div>
      </div></div>`;
      })
      .join("") || `<div class="hint">${t("noneYet")}</div>`
  }</div>`;
}
function widgetContent(w, s) {
  switch (w.type) {
    case "kpi":
      return renderKpiWidget(w, s);
    case "chart":
      return renderChartWidget(w);
    case "pie":
      return renderPieWidget(w);
    case "table":
      return renderTableWidget(w);
    case "breakdown":
      return renderBreakdownWidget(w, s);
    case "quota":
      return renderQuotaWidget(w);
    case "queue":
      return renderQueueWidget(w);
    case "rank":
      return renderRankWidget(w);
    default:
      return "";
  }
}
// layout default dipakai migrasi acara lama (v3.0) - mereproduksi tampilan
// dashboard tetap yg ada sebelumnya (KPI+tren+kuota+antrian+peringkat)
function defaultDashboardWidgets(cats) {
  // dulu pRank()/vBoard() selalu menampilkan 2 papan terpisah (pembeli
  // qty-tracked vs donatur non-qty) - pertahankan itu sbg 2 widget rank
  // yg masing2 dibatasi catIds-nya, bukan digabung jadi 1 daftar. catIds
  // kosong berarti "semua kategori di grup ini" di editor manual (lihat
  // widgetCatsInGroup), jadi kalau salah satu sisi (qty/flat) memang tidak
  // punya kategori sama sekali, widget itu dilewati saja - bukan dibikin
  // dengan catIds:[] yang malah kebaca "semua" dan jadi duplikat isinya.
  const qtyIds = cats.filter((c) => c.group === "income" && c.hasQty).map((c) => c.id);
  const flatIds = cats.filter((c) => c.group === "income" && !c.hasQty).map((c) => c.id);
  const widgets = [
    { id: uid(), type: "kpi", title: t("net"), metric: "net", x: 0, y: 0, w: 3, h: 1 },
    { id: uid(), type: "kpi", title: t("incomeW"), metric: "income", x: 3, y: 0, w: 3, h: 1 },
    { id: uid(), type: "kpi", title: t("exp"), metric: "expense", x: 6, y: 0, w: 3, h: 1 },
    { id: uid(), type: "kpi", title: t("wait"), metric: "pendingAmt", x: 9, y: 0, w: 3, h: 1 },
    { id: uid(), type: "chart", title: t("trend", { n: 14 }), x: 0, y: 1, w: 6, h: 4 },
    { id: uid(), type: "quota", title: t("seats"), catIds: [], x: 6, y: 1, w: 3, h: 4 },
    { id: uid(), type: "queue", title: t("queue"), catIds: [], x: 9, y: 1, w: 3, h: 4 },
  ];
  const rankW = qtyIds.length && flatIds.length ? 6 : 12;
  if (qtyIds.length)
    widgets.push({ id: uid(), type: "rank", title: t("topBuy"), group: "income", catIds: qtyIds, x: 0, y: 5, w: rankW, h: 3 });
  if (flatIds.length)
    widgets.push({
      id: uid(),
      type: "rank",
      title: t("topDon"),
      group: "income",
      catIds: flatIds,
      x: qtyIds.length ? 6 : 0,
      y: 5,
      w: rankW,
      h: 3,
    });
  return widgets;
}
// layout awal acara baru - cuma 3 KPI dasar, admin melengkapi sisanya lewat
// editor dashboard (Admin > Dashboard) setelah kategori acaranya didefinisikan
function starterDashboardWidgets() {
  return [
    { id: uid(), type: "kpi", title: t("incomeW"), metric: "income", x: 0, y: 0, w: 4, h: 1 },
    { id: uid(), type: "kpi", title: t("exp"), metric: "expense", x: 4, y: 0, w: 4, h: 1 },
    { id: uid(), type: "kpi", title: t("net"), metric: "net", x: 8, y: 0, w: 4, h: 1 },
  ];
}
function vDash() {
  const s = sums(),
    nw = narrow();
  const widgets = D().dashboard?.widgets || [];
  if (!widgets.length)
    return `<div class="fit" style="align-items:center;justify-content:center;text-align:center">
      <div class="empty">${t("noneYet")}${isAdmin() ? `<br><button class="btn sm" style="margin-top:10px" onclick="goDashEditor()">${t("editDashboard")}</button>` : ""}</div></div>`;
  const kpiWidgets = widgets.filter((w) => w.type === "kpi");
  const chartWidget = widgets.find((w) => w.type === "chart");
  const otherWidgets = widgets.filter((w) => w.type !== "kpi" && w.type !== "chart");
  if (nw) {
    const activeId = otherWidgets.some((w) => w.id === panel) ? panel : otherWidgets[0]?.id;
    const active = otherWidgets.find((w) => w.id === activeId);
    return `<div class="fit">
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(100px,1fr));flex:none">${kpiWidgets
      .map(
        (w) =>
          `<div class="card kpi act drill" style="padding:11px 13px;justify-content:center" onclick="drilldownKpi('${w.id}')">${renderKpiWidget(w, s)}</div>`,
      )
      .join("")}</div>
    ${chartWidget ? `<div class="card">${renderChartWidget(chartWidget)}</div>` : ""}
    ${
      otherWidgets.length
        ? `<div class="card" style="flex:1.1;gap:6px">
      <div class="seg" id="panelSeg" style="align-self:stretch;flex:none;max-width:100%">${otherWidgets
        .map(
          (w) =>
            `<button class="${activeId === w.id ? "on" : ""}" onclick="setPanel('${w.id}')">${esc(liveTitle(w.title))}</button>`,
        )
        .join("")}</div>
      <div style="display:flex;flex-direction:column;flex:1;min-height:0">${active ? widgetContent(active, s) : ""}</div>
    </div>`
        : ""
    }</div>`;
  }
  // sengaja BUKAN .fit (tinggi terbatas + scroll internal) - .fit ada di
  // dalam .wrap yg lebarnya dibatasi max-width & di-tengah-kan, jadi area
  // kosong di kiri/kanan layar lebar tidak akan pernah memicu scroll-nya.
  // Di sini dashboard mengalir sebagai konten halaman biasa supaya scroll
  // (mouse wheel/trackpad) bekerja di mana pun kursor berada di layar.
  return `<div style="padding:10px 0 100px">
    <div class="grid" style="grid-template-columns:repeat(${DASHBOARD_COLS},1fr);grid-auto-rows:${DASHBOARD_ROW_H}px;gap:12px">
      ${widgets
        .map((w) => {
          const kpi = w.type === "kpi";
          return `<div class="card${kpi ? " kpi act drill" : ""}" style="grid-column:${w.x + 1} / span ${Math.min(w.w, DASHBOARD_COLS)};grid-row:${w.y + 1} / span ${w.h};overflow:hidden;display:flex;flex-direction:column${kpi ? ";justify-content:center;padding:11px 13px" : ""}" ${kpi ? `onclick="drilldownKpi('${w.id}')"` : ""}>${widgetContent(w, s)}</div>`;
        })
        .join("")}
    </div>
  </div>`;
}
const rupiahTick = (v) =>
  v >= 1e6
    ? v / 1e6 + (lang === "id" ? " jt" : " M")
    : v >= 1000
      ? v / 1000 + (lang === "id" ? " rb" : " k")
      : v;
function drawChart() {
  const canvases = document.querySelectorAll("canvas[data-chart-widget]");
  const activeIds = new Set();
  const s = seriesByDay(narrow() ? 7 : 14);
  const cs = getComputedStyle(document.documentElement);
  const cv = (n) => cs.getPropertyValue(n).trim();
  canvases.forEach((el) => {
    const w = findWidget(el.dataset.chartWidget);
    const isLine = w?.chartStyle === "line";
    activeIds.add(el.id);
    chartInstances[el.id]?.destroy();
    chartInstances[el.id] = new Chart(el, {
      type: isLine ? "line" : "bar",
      data: {
        labels: s.map((x) => x.label),
        datasets: [
          {
            data: s.map((x) => x.v),
            backgroundColor: isLine ? cv("--blue-s") : cv("--blue"),
            hoverBackgroundColor: cv("--blue-d"),
            borderColor: cv("--blue"),
            borderWidth: isLine ? 2.5 : 0,
            borderRadius: isLine ? 0 : 5,
            maxBarThickness: 38,
            tension: 0.35,
            fill: isLine,
            pointRadius: isLine ? 3 : 0,
            pointBackgroundColor: cv("--blue"),
            pointHoverRadius: 5,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 500 },
        layout: { padding: { top: 6 } },
        onClick: (evt, elements) => {
          if (elements.length) drilldownDate(s[elements[0].index].k);
        },
        onHover: (evt, elements) => {
          evt.native.target.style.cursor = elements.length ? "pointer" : "default";
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (c) => rp(c.parsed.y),
              afterLabel: () => t("clickForDetail"),
            },
          },
        },
        scales: {
          x: {
            grid: { display: false },
            border: { display: false },
            ticks: { color: cv("--tx2"), font: { size: 12 } },
          },
          y: {
            grid: { color: cv("--line2") },
            border: { display: false },
            ticks: { color: cv("--tx2"), font: { size: 12 }, maxTicksLimit: 4, callback: rupiahTick },
          },
        },
      },
    });
  });
  Object.keys(chartInstances).forEach((id) => {
    if (!activeIds.has(id) && id.startsWith("wchart_")) {
      chartInstances[id]?.destroy();
      delete chartInstances[id];
    }
  });
}
function drawPies() {
  const canvases = document.querySelectorAll("canvas[data-pie-widget]");
  const activeIds = new Set();
  const s = sums();
  const cs = getComputedStyle(document.documentElement);
  const cv = (n) => cs.getPropertyValue(n).trim();
  canvases.forEach((el) => {
    const w = findWidget(el.dataset.pieWidget);
    if (!w) return;
    const cats = widgetCatsInGroup(w);
    const labels = cats.map((c) => c.n);
    const data = cats.map((c) => s.byCat[c.n] || 0);
    const total = data.reduce((a, b) => a + b, 0);
    const colors = cats.map((c, i) => cv(PIE_COLORS[i % PIE_COLORS.length]));
    activeIds.add(el.id);
    chartInstances[el.id]?.destroy();
    chartInstances[el.id] = new Chart(el, {
      type: "doughnut",
      data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: cv("--card"), borderWidth: 2 }] },
      plugins: [pieLabelsPlugin],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "62%",
        animation: { duration: 400 },
        onClick: (evt, elements) => {
          if (elements.length) drilldownCat(labels[elements[0].index]);
        },
        onHover: (evt, elements) => {
          evt.native.target.style.cursor = elements.length ? "pointer" : "default";
        },
        plugins: {
          legend: { position: "bottom", labels: { color: cv("--tx2"), boxWidth: 10, font: { size: 11 }, padding: 10 } },
          tooltip: {
            callbacks: {
              label: (c) => {
                const v = c.parsed,
                  pct = total ? Math.round((v / total) * 100) : 0;
                return `${c.label}: ${rp(v)} (${pct}%)`;
              },
              afterLabel: () => t("clickForDetail"),
            },
          },
        },
      },
    });
  });
  Object.keys(chartInstances).forEach((id) => {
    if (!activeIds.has(id) && id.startsWith("wpie_")) {
      chartInstances[id]?.destroy();
      delete chartInstances[id];
    }
  });
}

/* ================= dashboard (editor drag/resize, admin) ================= */
const WIDGET_TYPES = ["kpi", "chart", "pie", "table", "breakdown", "quota", "queue", "rank"];
const WIDGET_SIZE_DEFAULT = {
  kpi: [4, 1],
  chart: [6, 4],
  pie: [4, 4],
  table: [6, 4],
  breakdown: [4, 3],
  quota: [4, 4],
  queue: [4, 3],
  rank: [4, 3],
};
const widgetTypeLabel = (type) =>
  ({
    kpi: t("wKpi"),
    chart: t("wChart"),
    pie: t("wPie"),
    table: t("wTable"),
    breakdown: t("wBreakdown"),
    quota: t("wQuota"),
    queue: t("wQueue"),
    rank: t("wRank"),
  })[type] || type;
function vDashEditor() {
  const widgets = S.config.dashboard?.widgets || [];
  return `<div class="card" style="margin-bottom:12px">
    <div class="rowsp" style="flex-wrap:wrap;gap:8px">
      <p class="hint" style="margin:0;flex:1;min-width:220px">${t("dashEditorHint")}</p>
      <div style="display:flex;gap:8px;flex:none">
        <button class="btn ghost sm" onclick="openWidgetForm()">${t("addWidget")}</button>
        <button class="btn sm" onclick="saveDashLayout()">${t("saveLayout")}</button>
      </div>
    </div>
  </div>
  <div class="card" style="padding:10px;overflow-x:auto">
    ${
      widgets.length
        ? `<div class="grid-stack" id="dashGrid">${widgets
            .map(
              (w) => `<div class="grid-stack-item" gs-id="${w.id}" gs-x="${w.x}" gs-y="${w.y}" gs-w="${w.w}" gs-h="${w.h}">
        <div class="grid-stack-item-content dash-tile">
          <div class="dash-tile-head">
            <div style="min-width:0"><b style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:block">${esc(liveTitle(w.title))}</b>
            <span class="hint">${widgetTypeLabel(w.type)}</span></div>
            <div class="dash-tile-actions">
              <button type="button" class="btn ghost sm icon" onclick="openWidgetForm('${w.id}')" aria-label="${t("edit")}">✎</button>
              <button type="button" class="btn ghost sm icon" onclick="deleteDashWidget('${w.id}')" aria-label="${t("del")}">✕</button>
            </div>
          </div>
        </div>
      </div>`,
            )
            .join("")}</div>`
        : `<div class="empty">${t("noneYet")}</div>`
    }
  </div>`;
}
async function initDashGrid() {
  dashGrid?.destroy(false);
  dashGrid = null;
  const el = document.getElementById("dashGrid");
  if (!el) return;
  const GridStack = await loadGridstack();
  // .fit sudah pindah tab kalau admin sempat berpindah screen sebelum lib
  // ini selesai dimuat - jangan init grid pada elemen yang sudah lepas dari DOM
  if (!document.body.contains(el)) return;
  dashGrid = GridStack.init(
    { column: DASHBOARD_COLS, cellHeight: DASHBOARD_ROW_H, margin: 8, float: true },
    el,
  );
}
function openWidgetForm(id) {
  const w = id ? (S.config.dashboard?.widgets || []).find((x) => x.id === id) : null;
  const cats = D().cats;
  sheet(`<div class="rowsp" style="margin-bottom:14px"><h2 style="font-size:20px">${w ? t("edit") : t("addWidget")}</h2>${closeBtn()}</div>
    <input type="hidden" id="dw_id" value="${w ? w.id : ""}">
    <div class="field"><label>${t("widgetType")}</label>
      <select id="dw_type" onchange="onWidgetTypeChange()" ${w ? "disabled" : ""}>
        ${WIDGET_TYPES.map((ty) => `<option value="${ty}" ${w?.type === ty ? "selected" : ""}>${widgetTypeLabel(ty)}</option>`).join("")}
      </select></div>
    <div class="field"><label>${t("widgetTitleLbl")}</label><input id="dw_title" value="${esc(w?.title || "")}"></div>
    <div class="field" id="dw_style_wrap" style="display:none">
      <label>${t("chartStyleLbl")}</label>
      <select id="dw_style">
        <option value="bar" ${!w || w.chartStyle !== "line" ? "selected" : ""}>${t("chartStyleBar")}</option>
        <option value="line" ${w?.chartStyle === "line" ? "selected" : ""}>${t("chartStyleLine")}</option>
      </select></div>
    <div class="field" id="dw_metric_wrap" style="display:none">
      <label>${t("metricLbl")}</label>
      <select id="dw_metric">
        <option value="income" ${w?.metric === "income" ? "selected" : ""}>${t("incomeW")}</option>
        <option value="expense" ${w?.metric === "expense" ? "selected" : ""}>${t("exp")}</option>
        <option value="net" ${w?.metric === "net" ? "selected" : ""}>${t("net")}</option>
        <option value="pendingAmt" ${w?.metric === "pendingAmt" ? "selected" : ""}>${t("wait")}</option>
        ${cats.map((c) => `<option value="cat:${c.id}" ${w?.metric === "cat:" + c.id ? "selected" : ""}>${esc(c.n)}</option>`).join("")}
      </select></div>
    <div class="field" id="dw_group_wrap" style="display:none">
      <label>${t("groupLbl")}</label>
      <select id="dw_group" onchange="onWidgetGroupChange()">
        <option value="" ${w && !w.group ? "selected" : ""}>${t("allTypesLbl")}</option>
        <option value="income" ${!w || w.group === "income" ? "selected" : ""}>${t("incomeW")}</option>
        <option value="expense" ${w?.group === "expense" ? "selected" : ""}>${t("expW")}</option>
      </select></div>
    <div class="field" id="dw_cats_wrap" style="display:none">
      <label>${t("catFilterLbl")}</label>
      <div class="chips" id="dw_cats">${cats.map((c) => `<span class="chip ${w?.catIds?.includes(c.id) ? "on" : ""}" data-id="${c.id}" data-group="${c.group}" onclick="toggleEventChip(this)">${esc(c.n)}</span>`).join("") || `<span class="hint">${t("noneYet")}</span>`}</div>
      <div class="hint" style="margin-top:4px">${t("catFilterHint")}</div></div>
    <div class="rowsp" style="margin-top:14px">${w ? `<button class="btn danger" onclick="deleteDashWidget('${w.id}')">${t("del")}</button>` : "<span></span>"}
      <button class="btn" onclick="submitWidgetForm()">${t("save")}</button></div>`);
  onWidgetTypeChange();
}
const WIDGET_HAS_GROUP = ["breakdown", "pie", "table", "rank"];
const WIDGET_HAS_CATS = ["breakdown", "pie", "table", "quota", "rank"];
function onWidgetTypeChange() {
  const type = document.getElementById("dw_type").value;
  document.getElementById("dw_style_wrap").style.display = type === "chart" ? "block" : "none";
  document.getElementById("dw_metric_wrap").style.display = type === "kpi" ? "block" : "none";
  document.getElementById("dw_group_wrap").style.display = WIDGET_HAS_GROUP.includes(type)
    ? "block"
    : "none";
  document.getElementById("dw_cats_wrap").style.display = WIDGET_HAS_CATS.includes(type)
    ? "block"
    : "none";
  const titleEl = document.getElementById("dw_title");
  if (!titleEl.value.trim()) titleEl.value = widgetTypeLabel(type);
  onWidgetGroupChange();
}
// filter chip kategori yg tampil sesuai grup yg dipilih, biar admin tidak
// bisa (secara membingungkan) mencentang kategori expense pas grupnya income
function onWidgetGroupChange() {
  const group = document.getElementById("dw_group")?.value;
  document.querySelectorAll("#dw_cats .chip").forEach((el) => {
    el.style.display = !group || el.dataset.group === group ? "" : "none";
  });
}
async function submitWidgetForm() {
  const id = val("dw_id"),
    type = val("dw_type"),
    title = val("dw_title") || widgetTypeLabel(type);
  const metric = document.getElementById("dw_metric")?.value;
  const group = document.getElementById("dw_group")?.value;
  const chartStyle = document.getElementById("dw_style")?.value;
  const catIds = Array.from(document.querySelectorAll("#dw_cats .chip.on")).map(
    (el) => el.dataset.id,
  );
  await mutate(
    () => {
      if (!S.config.dashboard) S.config.dashboard = { widgets: [] };
      const widgets = S.config.dashboard.widgets;
      const patch = { title };
      if (type === "chart") patch.chartStyle = chartStyle;
      if (type === "kpi") patch.metric = metric;
      if (WIDGET_HAS_GROUP.includes(type)) patch.group = group || undefined;
      if (WIDGET_HAS_CATS.includes(type)) patch.catIds = catIds;
      const i = widgets.findIndex((w) => w.id === id);
      if (i >= 0) {
        Object.assign(widgets[i], patch);
      } else {
        const maxY = widgets.reduce((m, w) => Math.max(m, w.y + w.h), 0);
        const [w, h] = WIDGET_SIZE_DEFAULT[type] || [4, 3];
        widgets.push({ id: uid(), type, x: 0, y: maxY, w, h, ...patch });
      }
    },
    "actSet",
    title,
  );
  closeSheet();
  render();
}
async function deleteDashWidget(id) {
  await mutate(
    () => {
      S.config.dashboard.widgets = (S.config.dashboard.widgets || []).filter((w) => w.id !== id);
    },
    "actSet",
    t("del"),
  );
  closeSheet();
  render();
}
async function saveDashLayout() {
  if (!dashGrid) return toast(t("saveErr"));
  const positions = {};
  dashGrid.engine.nodes.forEach((n) => {
    positions[n.id] = { x: n.x, y: n.y, w: n.w, h: n.h };
  });
  await mutate(
    () => {
      (S.config.dashboard?.widgets || []).forEach((w) => {
        const p = positions[w.id];
        if (p) Object.assign(w, p);
      });
    },
    "actSet",
    t("saveLayout"),
  );
  toast(t("setSaved"));
}

/* ================= transaksi ================= */
function txMatches(x, q) {
  return (
    !q ||
    [x.name, x.phone, x.note, x.bank, x.cat, x.chequeNo, x.chequeBank, x.by]
      .join(" ")
      .toLowerCase()
      .includes(q)
  );
}
function sortTx(arr) {
  const { k, dir } = txSort,
    mul = dir === "asc" ? 1 : -1;
  return arr.slice().sort((a, b) => {
    let av = a[k],
      bv = b[k];
    if (k === "amount") {
      av = +av || 0;
      bv = +bv || 0;
    } else {
      av = String(av ?? "").toLowerCase();
      bv = String(bv ?? "").toLowerCase();
    }
    return av > bv ? mul : av < bv ? -mul : 0;
  });
}
function setTxSort(k) {
  txSort =
    txSort.k === k
      ? { k, dir: txSort.dir === "asc" ? "desc" : "asc" }
      : { k, dir: "asc" };
  txPage = 1;
  render();
}
// pencarian ketik-per-huruf TIDAK BOLEH memicu render() penuh - itu
// mengganti seluruh #root (termasuk <input> yg sedang difokus) dgn node DOM
// baru, yg di mobile langsung menutup keyboard tiap kali user mengetik satu
// huruf. Sebagai gantinya, cuma innerHTML div hasil daftarnya saja yg
// diperbarui (id="txListCard") - <input>-nya sendiri tidak pernah disentuh,
// jadi fokus & keyboard tetap terjaga. Pola yg sama dipakai di
// onLogSearch/onHubEventsSearch/onHubStaffSearch di bawah.
function onTxSearch(v) {
  txQ = v;
  txPage = 1;
  clearTimeout(window._tq);
  window._tq = setTimeout(() => {
    const el = document.getElementById("txListCard");
    if (el) el.innerHTML = txListCardHtml();
  }, 250);
}
// baris kartu ringkas (bukan <tr>) dipakai di layar sempit - tabel 5 kolom
// dipaksa muat 100% lebar viewport bikin tiap sel membungkus jadi berbaris2
// & kolom nominal/status kepotong (lihat txListCardHtml()). Gaya kartu ini
// pakai ulang class .drill-row yg sama dgn pop-up drill-down, spy konsisten.
function txRowCardHtml(x) {
  const isExp = x.type === "expense";
  const detail = [catLabel(x) + (x.seats ? " · " + x.seats : ""), x.bank || "-"]
    .filter(Boolean)
    .map(esc)
    .join(" · ");
  return `<div class="drill-row drill" onclick="openTx('${x.id}')">
    <div class="drill-row-bar ${isExp ? "drill-bar-neg" : "drill-bar-pos"}"></div>
    <div style="flex:1;min-width:0;display:flex;justify-content:space-between;gap:10px;align-items:flex-start">
      <div style="min-width:0">
        <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(x.name)}</div>
        <div class="hint" style="margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${detail}</div>
        <div class="hint" style="margin-top:1px">${x.date.slice(8)}/${x.date.slice(5, 7)}/${x.date.slice(0, 4)}</div>
      </div>
      <div style="text-align:right;flex:none;display:flex;flex-direction:column;align-items:flex-end;gap:6px">
        <div class="mono" style="font-weight:700${isExp ? ";color:var(--red)" : ""}">${isExp ? "−" : ""}${rp(x.amount)}</div>
        ${
          x.status === "pending"
            ? `<button class="btn ok sm" style="padding:5px 10px;font-size:12px" onclick="event.stopPropagation();verify('${x.id}')">${t("inRek")}</button>`
            : `<span class="tag t-ok">${t("inRek")}</span>`
        }
      </div>
    </div>
  </div>`;
}
function txListCardHtml() {
  const q = txQ.toLowerCase();
  const all = sortTx(
    S.tx.filter(
      (x) =>
        (filter === "all" ||
          (filter === "pending"
            ? x.status === "pending"
            : x.type === filter)) &&
        txMatches(x, q),
    ),
  );
  const { items: f, page, totalPages } = paginate(all, txPage);
  if (!f.length) return `<div class="empty"><b>${t("noTx")}</b><br>${t("noTxSub")}</div>`;
  if (narrow()) return f.map(txRowCardHtml).join("") + pagerHtml(page, totalPages, "setTxPage");
  const si = (k) =>
    txSort.k === k ? (txSort.dir === "asc" ? " ▲" : " ▼") : "";
  return `<div style="overflow-x:auto"><table><thead><tr>
    <th class="sortable" onclick="setTxSort('date')">${t("date")}${si("date")}</th>
    <th class="sortable" onclick="setTxSort('name')">${t("name")}${si("name")}</th>
    <th>${t("detail")}</th>
    <th class="sortable" style="text-align:right" onclick="setTxSort('amount')">${t("amount")}${si("amount")}</th>
    <th></th></tr></thead><tbody>
    ${f
      .map(
        (x) => `<tr>
      <td class="mono hint" style="white-space:nowrap">${x.date.slice(8)}/${x.date.slice(5, 7)}</td>
      <td><div style="font-weight:600">${esc(x.name)}</div><div class="hint">${esc(x.phone || "")}${x.by ? " · " + esc(x.by) : ""}</div></td>
      <td>${
        x.type === "expense"
          ? `<span class="tag t-exp">${t("expW")}</span> ${esc(catLabel(x))}`
          : `<span class="tag t-tic">${t("incomeW")}</span> ${esc(catLabel(x))}${x.seats ? " · " + x.seats : ""}`
      }
        <div class="hint">${esc(x.bank || "-")}${x.chequeNo ? " · " + t("chequeNo") + " " + esc(x.chequeNo) : ""}${x.note ? " · " + esc(x.note) : ""}</div></td>
      <td class="mono ${x.type === "expense" ? "neg" : ""}" style="text-align:right;white-space:nowrap;font-weight:600">${x.type === "expense" ? "−" : ""}${rp(x.amount)}
        <div style="margin-top:3px">${x.status === "verified" ? `<span class="tag t-ok">${t("inRek")}</span>` : `<span class="tag t-wait">${t("stW")}</span>`}</div></td>
      <td style="text-align:right;white-space:nowrap">${x.proof ? `<button class="btn ghost sm icon" onclick="viewProofById('${x.id}')" title="${t("viewProof")}">${proofIcon}</button> ` : ""}${x.status === "pending" ? `<button class="btn ok sm" onclick="verify('${x.id}')">${t("inRek")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openTx('${x.id}')">${t("edit")}</button></td></tr>`,
      )
      .join("")}
    </tbody></table></div>${pagerHtml(page, totalPages, "setTxPage")}`;
}
function vTx() {
  return `<div style="padding:12px 0 110px">
  <div class="rowsp" style="margin-bottom:11px;flex-wrap:wrap;gap:8px">
    <div class="chips">${[
      ["all", t("all")],
      ["pending", t("unchecked")],
      ["income", t("incomeW")],
      ["expense", t("expW")],
    ]
      .map(
        ([k, l]) =>
          `<button class="chip ${filter === k ? "on" : ""}" onclick="setFilter('${k}')">${l}</button>`,
      )
      .join("")}</div>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;min-width:200px">
      <input id="txSearchInput" style="max-width:280px" placeholder="${t("searchTx")}" value="${esc(txQ)}" oninput="onTxSearch(this.value)">
      <button class="btn ghost sm" onclick="openImport()">${t("imp")}</button></div></div>
  <div class="card" id="txListCard" style="padding:2px 10px">${txListCardHtml()}</div></div>`;
}
async function verify(id) {
  if (!canEdit()) return; // jaring pengaman - UI viewer seharusnya sudah tidak menampilkan tombol ini
  await mutate(
    () => {
      const x = S.tx.find((v) => v.id === id);
      if (x) {
        x.status = "verified";
        x.vBy = acting().name;
      }
    },
    "actVerify",
    (S.tx.find((v) => v.id === id) || {}).name +
      " · " +
      rp((S.tx.find((v) => v.id === id) || {}).amount),
  );
  toast(t("verified"));
  render();
}
function openTx(id) {
  const cats = D().cats;
  const x = id
    ? S.tx.find((v) => v.id === id)
    : {
        id: "",
        type: "income",
        date: today(),
        name: "",
        phone: "",
        cat: "",
        tier: "",
        seats: 0,
        amount: 0,
        bank: D().methods[0]?.name || "",
        bankName: "",
        note: "",
        proof: "",
        status: "pending",
      };
  if (!cats.length) {
    sheet(`<div class="rowsp" style="margin-bottom:14px"><h2 style="font-size:20px">${id ? t("editTx") : t("newTx")}</h2>${closeBtn()}</div>
      <div class="empty">${t("noCatYet")}</div>`);
    return;
  }
  const initialGroup = x.type === "expense" ? "expense" : "income";
  // viewer: tampilkan detail transaksi tanpa kemampuan mengubah apa pun -
  // seluruh isi form dibungkus pointer-events:none (satu titik kontrol,
  // bukan menandai "disabled" di puluhan field satu per satu), kecuali
  // gambar bukti transfer yg tetap boleh diklik utk dilihat (aksi baca)
  const readOnly = !canEdit();
  sheet(`<div class="rowsp" style="margin-bottom:14px"><h2 style="font-size:20px">${readOnly ? t("txDetail") : id ? t("editTx") : t("newTx")}</h2>
      ${closeBtn()}</div>
    <div id="txFormBody" style="${readOnly ? "pointer-events:none;opacity:.7" : ""}">
    <input type="hidden" id="f_id" value="${x.id}"><input type="hidden" id="f_type" value="${x.type}">
    <div class="seg" id="txTabs" style="margin-bottom:14px">
      <button type="button" class="${initialGroup === "income" ? "on" : ""}" data-g="income" onclick="setTxTab('income')">${t("incomeW")}</button>
      <button type="button" class="${initialGroup === "expense" ? "on" : ""}" data-g="expense" onclick="setTxTab('expense')">${t("expW")}</button>
    </div>
    <div class="field" style="margin-bottom:14px">
      <label id="l_cat">${t("txCat")} <span class="req">*</span></label>
      <select id="f_cat" onchange="onCatChange()"></select>
    </div>
    <div class="field" id="w_tier" style="display:none;margin-bottom:14px">
      <label>${t("tierLbl")} <span class="req">*</span></label>
      <select id="f_tier" onchange="onTierChange()"></select>
    </div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field"><label>${t("date")} <span class="req">*</span></label><input type="date" id="f_date" value="${x.date}"></div>
      <div class="field"><label id="l_bank">${t("paymentMethod")} <span class="req">*</span></label><select id="f_bank" onchange="onMethodChange()">${D()
        .methods.map(
          (m) =>
            `<option value="${esc(m.name)}" data-type="${m.type}" ${m.name === x.bank ? "selected" : ""}>${esc(m.name)}</option>`,
        )
        .join("")}</select></div>
      <div class="field" id="w_bankName" style="display:none">
        <label>${t("bankNameLbl")} <span class="req">*</span></label>
        <select id="f_bankName">${INDONESIA_BANKS.map((b) => `<option ${b === x.bankName ? "selected" : ""}>${esc(b)}</option>`).join("")}</select></div>
      <div class="field" id="w_name"><label id="l_name">${t("nameGeneric")} <span class="req">*</span></label><input id="f_name" value="${esc(x.name)}"></div>
      <div class="field" id="w_phone"><label>${t("wa")}</label><input id="f_phone" inputmode="tel" value="${esc(x.phone)}" placeholder="0812..."></div>
      <div class="field" id="w_seats"><label>${t("nSeat")}</label>
        <div class="stepper">
          <button type="button" class="step-btn" onclick="stepSeats(-1)" aria-label="-">−</button>
          <input type="number" inputmode="numeric" min="0" id="f_seats" value="${x.seats || 0}" oninput="recalc()">
          <button type="button" class="step-btn" onclick="stepSeats(1)" aria-label="+">+</button>
        </div></div>
    </div>
    <div class="grid" id="w_cheque" style="display:none;grid-template-columns:1fr 1fr 1fr">
      <div class="field"><label>${t("chequeNo")}</label><input id="f_chequeNo" value="${esc(x.chequeNo || "")}"></div>
      <div class="field"><label>${t("chequeBank")}</label><input id="f_chequeBank" value="${esc(x.chequeBank || "")}"></div>
      <div class="field"><label>${t("chequeDate")}</label><input type="date" id="f_chequeDate" value="${x.chequeDate || ""}"></div>
    </div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field amt-field"><label id="l_amount">${t("amtIn")} <span class="req">*</span></label><input type="number" inputmode="numeric" id="f_amount" value="${x.amount}" oninput="prev();updateBonusHint()">
        <div class="hint mono" id="prev" style="margin-top:4px">${rp(x.amount)}</div></div>
      <div class="field"><label>${t("payStatus")}</label><select id="f_status">
        <option value="pending" ${x.status === "pending" ? "selected" : ""}>${t("stW")}</option>
        <option value="verified" ${x.status === "verified" ? "selected" : ""}>${t("stV")}</option></select></div></div>
    <div id="w_bonusHint" style="display:none;margin-bottom:14px"></div>
    <div class="field"><label>${t("note")}</label><textarea id="f_note" rows="2">${esc(x.note)}</textarea></div>
    <div class="field"><label>${t("note2")}</label><textarea id="f_note2" rows="2">${esc(x.note2 || "")}</textarea></div>
    <div class="field"><label>${t("proof")}</label>
      <div class="filepick">
        <button type="button" class="btn ghost sm" onclick="document.getElementById('f_proofFile').click()">${t("chooseFile")}</button>
        <span class="hint" id="proofFileName">${x.proof ? t("fileAttached") : t("noFileChosen")}</span>
      </div>
      <input type="file" id="f_proofFile" accept="image/*" style="display:none" onchange="onProofChange(event)">
      <div class="hint" style="margin-top:4px">${t("proofHint")}</div>
      <div class="proof-prev" id="proofPrev" style="margin-top:8px${readOnly ? ";pointer-events:auto" : ""}">${x.proof ? `<img src="${x.proof}" onclick="viewProofPreview()">` : ""}</div>
      <input type="hidden" id="f_proof" value="${esc(x.proof || "")}"></div>
    ${(() => {
      const cf = D().tpl.filter((c) => c.custom);
      if (!cf.length) return "";
      return `<div class="field"><label>${t("customFields")}</label>
        <div class="grid" style="grid-template-columns:1fr 1fr">${cf
          .map(
            (c) =>
              `<div class="field"><label>${esc(c.h)}</label><input id="fc_${c.f}" value="${esc((x.custom && x.custom[c.f]) || "")}"></div>`,
          )
          .join("")}</div></div>`;
    })()}
    </div>
    ${
      readOnly
        ? ""
        : `<div class="rowsp" style="margin-top:14px">${id ? `<button class="btn danger" onclick="delTx('${id}')">${t("del")}</button>` : "<span></span>"}
      <button class="btn" onclick="saveTx()">${t("save")}</button></div>`
    }`);
  setTxTab(initialGroup, x);
  if (!readOnly) {
    initCombobox("f_cat");
    initCombobox("f_bankName");
  }
}
function sheet(html) {
  const m = document.createElement("div");
  m.className = "modal";
  m.id = "modal";
  m.onclick = (e) => {
    if (e.target === m) m.remove();
  };
  m.innerHTML = `<div class="sheet">${html}</div>`;
  document.body.appendChild(m);
}
function closeSheet() {
  document.getElementById("modal")?.remove();
}
function prev() {
  document.getElementById("prev").textContent = rp(
    +document.getElementById("f_amount").value || 0,
  );
}
function compressImage(file, maxDim, quality) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      let w = img.naturalWidth,
        h = img.naturalHeight;
      if (w > h && w > maxDim) {
        h = Math.round((h * maxDim) / w);
        w = maxDim;
      } else if (h > maxDim) {
        w = Math.round((w * maxDim) / h);
        h = maxDim;
      }
      const c = document.createElement("canvas");
      c.width = w;
      c.height = h;
      c.getContext("2d").drawImage(img, 0, 0, w, h);
      resolve(c.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("bad image"));
    };
    img.src = url;
  });
}
async function onProofChange(e) {
  const file = e.target.files[0];
  if (!file) return;
  if (file.size > 8 * 1024 * 1024) {
    toast(t("proofBig"));
    e.target.value = "";
    return;
  }
  try {
    const dataUrl = await compressImage(file, 1280, 0.72);
    document.getElementById("f_proof").value = dataUrl;
    document.getElementById("proofFileName").textContent = file.name;
    document.getElementById("proofPrev").innerHTML =
      `<img src="${dataUrl}" onclick="viewProofPreview()">`;
  } catch (err) {
    toast(t("proofBig"));
    e.target.value = "";
  }
}
function viewProof(src) {
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${t("viewProof")}</h2>${closeBtn()}</div>
    <img src="${src}" style="max-width:100%;border-radius:10px;display:block">`);
}
function viewProofPreview() {
  const v = document.getElementById("f_proof")?.value;
  if (v) viewProof(v);
}
function viewProofById(id) {
  const x = S.tx.find((v) => v.id === id);
  if (x && x.proof) viewProof(x.proof);
}
// tab Pemasukan/Pengeluaran di atas dropdown kategori - membatasi pilihan
// kategori ke grup yg aktif, jadi admin tak perlu menyisir daftar gabungan
function catOptionHtml(c) {
  const tiersJson = esc(JSON.stringify(c.tiers || [])).replace(/'/g, "&#39;");
  return `<option value="${esc(c.n)}" data-qty="${c.hasQty ? 1 : 0}" data-price="${c.p || 0}" data-tiers='${tiersJson}'>${esc(c.n)}</option>`;
}
function setTxTab(group, x) {
  document
    .querySelectorAll("#txTabs button")
    .forEach((b) => b.classList.toggle("on", b.dataset.g === group));
  const cats = D().cats.filter((c) => c.group === group);
  const sel = document.getElementById("f_cat");
  sel.innerHTML = cats.length
    ? cats.map(catOptionHtml).join("")
    : `<option value="" disabled>${t("noCatYet")}</option>`;
  if (x && cats.some((c) => c.n === x.cat)) sel.value = x.cat;
  syncCombo("f_cat");
  onCatChange(true, x);
}
// dipanggil tiap kali kategori transaksi diganti - grup (income/expense)
// sekarang datang dari tab yg aktif, bukan dari option itu sendiri; apakah
// kategori itu "melacak jumlah" & daftar tingkat harganya (mis. Pembelian
// Kursi -> Platinum/Gold/Silver) ada di data-attribute option, lihat
// catOptionHtml()
function onCatChange(init, x) {
  const opt = document.getElementById("f_cat")?.selectedOptions[0];
  const group = document.querySelector("#txTabs button.on")?.dataset.g || "income",
    hasQty = opt?.dataset.qty === "1",
    tiers = opt?.dataset.tiers ? JSON.parse(opt.dataset.tiers) : [];
  document.getElementById("f_type").value = group;
  const e = group === "expense";
  const tierWrap = document.getElementById("w_tier"),
    tierSel = document.getElementById("f_tier");
  tierWrap.style.display = tiers.length ? "block" : "none";
  if (tiers.length) {
    tierSel.innerHTML = tiers
      .map((tr) => `<option value="${esc(tr.name)}" data-price="${tr.p}">${esc(tr.name)} (${rp(tr.p)})</option>`)
      .join("");
    const want = init && x?.tier ? x.tier : tierSel.value;
    if (tiers.some((tr) => tr.name === want)) tierSel.value = want;
  }
  document.getElementById("w_seats").style.display = hasQty ? "block" : "none";
  document.getElementById("w_phone").style.display = e ? "none" : "block";
  // pengeluaran menyembunyikan phone, jadi field nama jadi sendirian di
  // barisnya - bentangkan penuh 2 kolom biar tidak kelihatan terpotong separuh
  document.getElementById("w_name").style.gridColumn = e ? "1 / -1" : "";
  document.getElementById("l_name").innerHTML = `${e ? t("namePay") : t("nameGeneric")} <span class="req">*</span>`;
  if (!init) {
    if (hasQty) recalc();
    else {
      const price = tiers.length ? +tiers[0]?.p || 0 : +(opt?.dataset.price || 0);
      if (price) {
        document.getElementById("f_amount").value = price;
        prev();
      }
    }
  }
  onMethodChange();
  updateBonusHint();
}
function onTierChange() {
  recalc();
}
// hint informasional (bukan otomatis mencatat transaksi baru) - kalau
// kategori yg dipilih (mis. Sponsor) punya bonusRules dan nominalnya
// memenuhi syarat, tampilkan apa yg berhak didapat penyumbangnya
function updateBonusHint() {
  const hintEl = document.getElementById("w_bonusHint");
  const catOpt = document.getElementById("f_cat")?.selectedOptions[0];
  const cat = D().cats.find((c) => c.n === catOpt?.value);
  if (!cat?.bonusRules?.length) {
    hintEl.style.display = "none";
    return;
  }
  const amt = +document.getElementById("f_amount")?.value || 0;
  const applicable = cat.bonusRules
    .filter((r) => amt >= r.minAmount)
    .sort((a, b) => b.minAmount - a.minAmount)[0];
  if (!applicable) {
    hintEl.style.display = "none";
    return;
  }
  const targetCat = D().cats.find((c) => c.id === applicable.targetCatId);
  hintEl.style.display = "block";
  hintEl.innerHTML = `<div class="hint" style="padding:10px 12px;background:var(--blue-s);border-radius:9px;color:var(--blue)">${t("bonusHint", { qty: applicable.freeQty, tier: applicable.targetTier, cat: targetCat?.n || "" })}</div>`;
}
function stepSeats(d) {
  const el = document.getElementById("f_seats");
  el.value = Math.max(0, (+el.value || 0) + d);
  recalc();
}
function recalc() {
  const opt = document.getElementById("f_cat")?.selectedOptions[0];
  if (!opt || opt.dataset.qty !== "1") return;
  const tierVisible = document.getElementById("w_tier").style.display !== "none";
  const price = tierVisible
    ? +(document.getElementById("f_tier")?.selectedOptions[0]?.dataset.price || 0)
    : +(opt.dataset.price || 0);
  document.getElementById("f_amount").value = (+document.getElementById("f_seats").value || 0) * price;
  prev();
}
function currentMethodType() {
  const sel = document.getElementById("f_bank");
  return sel?.selectedOptions[0]?.dataset.type || "bank";
}
// dropdown panjang (kategori, daftar bank, acara) susah dicari manual - bukan
// search-box + <select> terpisah (2 elemen), tapi SATU field gabungan:
// <select> aslinya disembunyikan (tetap dipakai sbg sumber kebenaran, semua
// logika data-attribute/onchange yg sudah ada tidak berubah), lalu di
// depannya dipasang <input> yg sekaligus jadi tampilan nilai terpilih DAN
// kotak pencarian - fokus/ketik membuka daftar opsi yg cocok, klik salah
// satu men-set select-nya lalu memicu event "change" spt biasa
function initCombobox(selectId, placeholder) {
  const sel = document.getElementById(selectId);
  if (!sel || sel.dataset.comboInit) return;
  sel.dataset.comboInit = "1";
  sel.classList.add("combo-native");
  const wrap = document.createElement("div");
  wrap.className = "combobox";
  const input = document.createElement("input");
  input.type = "text";
  input.className = "combo-input";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.placeholder = placeholder || t("searchPlaceholder");
  const list = document.createElement("div");
  list.className = "combo-list";
  const currentLabel = () => {
    const o = sel.selectedOptions[0];
    return o && o.value !== "" ? o.textContent : "";
  };
  const closeList = () => list.classList.remove("open");
  const renderList = (filterText) => {
    const q = (filterText || "").trim().toLowerCase();
    const opts = Array.from(sel.options).filter((o) => !o.disabled);
    const filtered = opts.filter((o) => !q || o.textContent.toLowerCase().includes(q));
    list.innerHTML = filtered.length
      ? filtered
          .map(
            (o) =>
              `<div class="combo-opt${o.value === sel.value ? " sel" : ""}" data-value="${esc(o.value)}">${esc(o.textContent)}</div>`,
          )
          .join("")
      : `<div class="combo-empty">${t("noneYet")}</div>`;
  };
  input.addEventListener("focus", () => {
    input.select();
    renderList("");
    list.classList.add("open");
  });
  input.addEventListener("input", () => {
    renderList(input.value);
    list.classList.add("open");
  });
  list.addEventListener("mousedown", (e) => {
    const optEl = e.target.closest(".combo-opt");
    if (!optEl) return;
    e.preventDefault();
    sel.value = optEl.dataset.value;
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    input.value = currentLabel();
    closeList();
  });
  input.addEventListener("blur", () => {
    setTimeout(() => {
      input.value = currentLabel();
      closeList();
    }, 150);
  });
  wrap.appendChild(input);
  wrap.appendChild(list);
  sel.parentNode.insertBefore(wrap, sel);
  input.value = currentLabel();
  // fungsi lain (mis. setTxTab saat ganti tab income/expense) mengubah
  // value/opsi select scr programatik - panggil syncCombo(selectId) sesudahnya
  // supaya teks yg tampil di field ikut ter-update
  sel._comboSync = () => {
    input.value = currentLabel();
  };
}
function syncCombo(selectId) {
  document.getElementById(selectId)?._comboSync?.();
}
// label field nominal harus sesuai metode pembayaran yg dipilih - "Nominal
// transfer" tidak masuk akal utk Tunai/Utang/Cek, jadi labelnya menyesuaikan
function amountLabelFor(type, isExpense) {
  if (type === "cash") return t("amtCash");
  if (type === "cheque") return t("amtCheque");
  if (type === "debt") return t("amtDebt");
  if (type === "other") return t("amount");
  return isExpense ? t("amtOut") : t("amtIn");
}
function onMethodChange() {
  const type = currentMethodType();
  document.getElementById("w_cheque").style.display = type === "cheque" ? "grid" : "none";
  document.getElementById("w_bankName").style.display = type === "bank" ? "block" : "none";
  const isExpense = document.querySelector("#txTabs button.on")?.dataset.g === "expense";
  document.getElementById("l_amount").innerHTML = `${amountLabelFor(type, isExpense)} <span class="req">*</span>`;
  // utang belum ada uang yg benar2 diterima/dibayarkan - default status ke
  // "belum lunas" tiap kali metode diganti ke utang (masih bisa diubah manual
  // nanti kalau utangnya sudah dilunasi)
  if (type === "debt") {
    const st = document.getElementById("f_status");
    if (st) st.value = "pending";
  }
}
// kolom bertanda * merah di form wajib diisi - cek di sini sebelum simpan,
// bukan diam2 dikasih nilai default (dulu nama kosong jadi "—") supaya data
// yg tersimpan memang benar2 sudah diisi user, bukan hasil tebakan aplikasi
function invalidTxFields() {
  const g = (i) => document.getElementById("f_" + i)?.value || "";
  const methodType = currentMethodType();
  const catOpt = document.getElementById("f_cat")?.selectedOptions[0];
  const tiers = catOpt?.dataset.tiers ? JSON.parse(catOpt.dataset.tiers) : [];
  const bad = [];
  if (!g("date")) bad.push("f_date");
  if (!g("cat")) bad.push("f_cat");
  if (tiers.length && !g("tier")) bad.push("f_tier");
  if (!g("name").trim()) bad.push("f_name");
  if (!g("bank")) bad.push("f_bank");
  if (methodType === "bank" && !g("bankName")) bad.push("f_bankName");
  if (!(+g("amount") > 0)) bad.push("f_amount");
  return bad;
}
async function saveTx() {
  if (!canEdit()) return; // jaring pengaman - UI viewer seharusnya sudah read-only
  document.querySelectorAll("#modal .field.invalid").forEach((el) => el.classList.remove("invalid"));
  const invalid = invalidTxFields();
  if (invalid.length) {
    invalid.forEach((fieldId) => document.getElementById(fieldId)?.closest(".field")?.classList.add("invalid"));
    toast(t("fillAll"));
    document.getElementById(invalid[0])?.focus();
    return;
  }
  const g = (i) => document.getElementById("f_" + i).value,
    id = g("id"),
    methodType = currentMethodType(),
    catOpt = document.getElementById("f_cat")?.selectedOptions[0],
    ty = document.querySelector("#txTabs button.on")?.dataset.g || "income",
    hasQty = catOpt?.dataset.qty === "1",
    tiers = catOpt?.dataset.tiers ? JSON.parse(catOpt.dataset.tiers) : [];
  const x = {
    id: id || uid(),
    type: ty,
    date: g("date") || today(),
    name: g("name").trim(),
    phone: ty === "expense" ? "" : g("phone"),
    cat: g("cat"),
    tier: tiers.length ? g("tier") : "",
    seats: hasQty ? +g("seats") || 0 : 0,
    amount: +g("amount") || 0,
    bank: g("bank"),
    bankName: methodType === "bank" ? g("bankName") : "",
    methodType,
    chequeNo: methodType === "cheque" ? g("chequeNo") : "",
    chequeBank: methodType === "cheque" ? g("chequeBank") : "",
    chequeDate: methodType === "cheque" ? g("chequeDate") : "",
    note: g("note"),
    note2: g("note2"),
    proof: g("proof"),
    status: g("status"),
    by: acting().name,
    custom: Object.fromEntries(
      D()
        .tpl.filter((c) => c.custom)
        .map((c) => [
          c.f,
          (document.getElementById("fc_" + c.f)?.value || "").trim(),
        ]),
    ),
  };
  await mutate(
    () => {
      const i = S.tx.findIndex((v) => v.id === x.id);
      i >= 0 ? (S.tx[i] = x) : S.tx.push(x);
    },
    id ? "actUpdate" : "actCreate",
    `${x.name} · ${rp(x.amount)} · ${x.status === "verified" ? t("inRek") : t("stW")}`,
  );
  closeSheet();
  toast(t("saved"));
  render();
}
async function delTx(id) {
  if (!canEdit()) return; // jaring pengaman - UI viewer seharusnya sudah read-only
  if (!confirm(t("confirmDel"))) return;
  const x = S.tx.find((v) => v.id === id) || {};
  await mutate(
    () => {
      S.tx = S.tx.filter((v) => v.id !== id);
    },
    "actDelete",
    `${x.name} · ${rp(x.amount)}`,
  );
  closeSheet();
  toast(t("deleted"));
  render();
}

/* ================= peringkat ================= */
function vBoard() {
  const BOARD_SIZE = 10;
  const qtyNames = D()
      .cats.filter((c) => c.group === "income" && c.hasQty)
      .map((c) => c.n),
    flatNames = D()
      .cats.filter((c) => c.group === "income" && !c.hasQty)
      .map((c) => c.n);
  const allBuy = byPerson(qtyNames),
    allDon = byPerson(flatNames);
  const { items: a, page: pageBuy, totalPages: totalBuy } = paginate(allBuy, boardBuyPage, BOARD_SIZE);
  const { items: b, page: pageDon, totalPages: totalDon } = paginate(allDon, boardDonPage, BOARD_SIZE);
  const maxBuy = allBuy.reduce((m, x) => Math.max(m, x.amount), 0) || 1;
  const maxDon = allDon.reduce((m, x) => Math.max(m, x.amount), 0) || 1;
  // dipakai medal + bar proporsi yg sama dgn renderRankWidget() di dashboard,
  // supaya bahasa visualnya konsisten di seluruh app - lihat MEDAL_COLORS
  const row = (x, rank, max) => {
    const pct = Math.max(6, Math.round((x.amount / max) * 100));
    const medal = rank <= 3 ? MEDAL_COLORS[rank - 1] : null;
    const detail = `${x.n} ${t("trans")}${x.seats ? " · " + x.seats + " " + t("seatsW").toLowerCase() : ""}`;
    return `<div class="rank-row" style="animation:slidein .3s var(--ease) backwards;animation-delay:${((rank - 1) % BOARD_SIZE) * 35}ms" title="${esc(detail)}">
      <div class="rank-badge" style="${medal ? `background:${medal};color:#fff` : "background:var(--line2);color:var(--tx2)"}">${rank}</div>
      <div style="flex:1;min-width:0">
        <div class="rowsp" style="gap:8px"><span class="rank-name">${esc(x.name)}</span><span class="mono rank-amt">${rp(x.amount)}</span></div>
        <div class="rank-bar"><i style="width:${pct}%;background:${medal || "var(--blue)"}"></i></div>
      </div></div>`;
  };
  return `<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(330px,1fr));padding:12px 0 110px">
    <div class="card"><h2 style="font-size:17px;margin-bottom:8px">${t("topBuy")}</h2>
      ${a.map((x, i) => row(x, (pageBuy - 1) * BOARD_SIZE + i + 1, maxBuy)).join("") || `<div class="empty">${t("noneYet")}</div>`}
      ${pagerHtml(pageBuy, totalBuy, "setBoardBuyPage")}</div>
    <div class="card"><h2 style="font-size:17px;margin-bottom:8px">${t("topDon")}</h2>
      ${b.map((x, i) => row(x, (pageDon - 1) * BOARD_SIZE + i + 1, maxDon)).join("") || `<div class="empty">${t("noneYet")}</div>`}
      ${pagerHtml(pageDon, totalDon, "setBoardDonPage")}</div></div>`;
}

/* ================= admin ================= */
function vAdmin() {
  if (!isAdmin()) return `<div class="empty">${t("onlyAdmin")}</div>`;
  return `<div style="padding:12px 0 110px">
    <div class="seg" style="margin-bottom:12px">${[
      ["logs", t("logs")],
      ["set", t("set")],
      ["dashboard", t("dashboardTab")],
    ]
      .map(
        ([k, l]) =>
          `<button class="${adminTab === k ? "on" : ""}" onclick="setAdminTab('${k}')">${l}</button>`,
      )
      .join("")}</div>
    ${{ logs: vLogs, set: vSet, dashboard: vDashEditor }[adminTab]()}</div>`;
}
function hubStaffListHtml() {
  const q = hubStaffQ.toLowerCase();
  let all = G.users.filter(
    (u) =>
      (hubStaffFilter === "all" || u.role === hubStaffFilter) &&
      (!q || (u.name + " " + u.email).toLowerCase().includes(q)),
  );
  const { k, dir } = hubStaffSort,
    mul = dir === "asc" ? 1 : -1;
  all = all.slice().sort((a, b) => {
    const av = String(a[k] ?? "").toLowerCase(),
      bv = String(b[k] ?? "").toLowerCase();
    return av > bv ? mul : av < bv ? -mul : 0;
  });
  const { items, page, totalPages } = paginate(all, hubStaffPage, 10);
  if (!items.length) return `<div class="empty">${t("noneYet")}</div>`;
  if (narrow())
    return (
      items
        .map(
          (u) => `<div class="drill-row">
      <div class="drill-row-bar" style="background:${u.active ? "var(--green)" : "var(--red)"}"></div>
      <div style="flex:1;min-width:0">
        <div class="rowsp" style="gap:8px;align-items:flex-start">
          <div style="min-width:0">
            <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(u.name)}</div>
            <div class="hint" style="margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(u.email)}${u.provider === "google" ? " · Google" : ""}</div>
          </div>
          <div style="text-align:right;flex:none">
            <span class="tag ${roleTagClass(u.role)}">${roleLabel(u.role)}</span>
            <div class="hint" style="margin-top:4px">${u.active ? t("active") : t("inactive")}</div>
          </div>
        </div>
        <div class="hint mono" style="margin-top:4px;font-size:12px">${t("lastIn")}: ${u.last ? new Date(u.last).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : t("never")}</div>
        <div class="rowsp" style="margin-top:8px;gap:8px">
          ${u.email !== me.email && u.active ? `<button class="btn ghost sm" onclick="startImp('${esc(u.email)}')">${t("loginAs")}</button>` : "<span></span>"}
          <button class="btn ghost sm" onclick="openUser('${esc(u.email)}')">${t("edit")}</button>
        </div>
      </div></div>`,
        )
        .join("") + pagerHtml(page, totalPages, "setHubStaffPage")
    );
  const si = (k) =>
    hubStaffSort.k === k ? (hubStaffSort.dir === "asc" ? " ▲" : " ▼") : "";
  return `<div style="overflow-x:auto"><table><thead><tr>
    <th class="sortable" onclick="setHubStaffSort('name')">${t("name")}${si("name")}</th>
    <th class="sortable" onclick="setHubStaffSort('role')">${t("role")}${si("role")}</th>
    <th class="sortable" onclick="setHubStaffSort('active')">${t("status")}${si("active")}</th>
    <th class="sortable" onclick="setHubStaffSort('last')">${t("lastIn")}${si("last")}</th><th></th></tr></thead><tbody>
    ${items
      .map(
        (u) => `<tr>
      <td><div style="font-weight:600">${esc(u.name)}</div><div class="hint">${esc(u.email)}${u.provider === "google" ? " · Google" : ""}</div></td>
      <td><span class="tag ${roleTagClass(u.role)}">${roleLabel(u.role)}</span></td>
      <td><span class="tag ${u.active ? "t-ok" : "t-exp"}">${u.active ? t("active") : t("inactive")}</span></td>
      <td class="hint mono">${u.last ? new Date(u.last).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : t("never")}</td>
      <td style="text-align:right;white-space:nowrap">
        ${u.email !== me.email && u.active ? `<button class="btn ghost sm" onclick="startImp('${esc(u.email)}')">${t("loginAs")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openUser('${esc(u.email)}')">${t("edit")}</button></td></tr>`,
      )
      .join("")}
  </tbody></table></div>${pagerHtml(page, totalPages, "setHubStaffPage")}`;
}
function vHubStaff() {
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("staff")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;flex-wrap:wrap;min-width:220px">
      <div class="chips">${[
        ["all", t("all")],
        ["admin", t("admins")],
        ["treasurer", t("treas")],
        ["viewer", t("viewer")],
      ]
        .map(
          ([k2, l]) =>
            `<button class="chip ${hubStaffFilter === k2 ? "on" : ""}" onclick="setHubStaffFilter('${k2}')">${l}</button>`,
        )
        .join("")}</div>
      <input id="hubStaffSearchInput" style="max-width:220px" placeholder="${t("searchTx")}" value="${esc(hubStaffQ)}" oninput="onHubStaffSearch(this.value)">
      <button class="btn sm" onclick="openUser()">${t("addUser")}</button></div></div>
  <div id="hubStaffListBody">${hubStaffListHtml()}</div></div>`;
}
function setHubStaffFilter(k) {
  hubStaffFilter = k;
  hubStaffPage = 1;
  render();
}
function toggleEventChip(el) {
  el.classList.toggle("on");
}
function openUser(email) {
  const u = email
    ? G.users.find((x) => x.email === email)
    : {
        name: "",
        email: "",
        role: "treasurer",
        active: true,
        provider: "password",
        eventIds: [],
      };
  const activeEvents = G.events.filter((e) => e.status === "active");
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${email ? t("edit") : t("addUser")}</h2>
    ${closeBtn()}</div>
    <input type="hidden" id="v_old" value="${esc(email || "")}">
    <div class="field"><label>${t("fullName")}</label><input id="v_name" value="${esc(u.name)}"></div>
    <div class="field"><label>${t("email")}</label><input id="v_email" type="email" value="${esc(u.email)}" ${email ? "disabled" : ""}></div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field"><label>${t("role")}</label><select id="v_role" onchange="document.getElementById('v_events_wrap').style.display=this.value==='admin'?'none':'block'">
        <option value="treasurer" ${u.role === "treasurer" ? "selected" : ""}>${t("treas")}</option>
        <option value="viewer" ${u.role === "viewer" ? "selected" : ""}>${t("viewer")}</option>
        <option value="admin" ${u.role === "admin" ? "selected" : ""}>${t("admins")}</option></select></div>
      <div class="field"><label>${t("status")}</label><select id="v_active">
        <option value="1" ${u.active ? "selected" : ""}>${t("active")}</option>
        <option value="0" ${!u.active ? "selected" : ""}>${t("inactive")}</option></select></div></div>
    <div class="field"><label>${email ? t("newPass") : t("pass")}</label><input id="v_pass" type="password" placeholder="${email ? "—" : ""}"></div>
    <div class="field" id="v_events_wrap" style="${u.role === "admin" ? "display:none" : ""}">
      <label>${t("events")}</label>
      <div class="chips" id="v_events">${
        activeEvents.length
          ? activeEvents
              .map(
                (e) =>
                  `<span class="chip ${(u.eventIds || []).includes(e.id) ? "on" : ""}" data-id="${e.id}" onclick="toggleEventChip(this)">${esc(e.name)}</span>`,
              )
              .join("")
          : `<span class="hint">${t("noEventsAssigned")}</span>`
      }</div>
    </div>
    <div class="rowsp" style="margin-top:12px">
      ${email && email !== me.email ? `<button class="btn danger" onclick="delUser('${esc(email)}')">${t("del")}</button>` : "<span></span>"}
      <button class="btn" onclick="saveUser()">${t("saveGeneric")}</button></div>`);
}
async function saveUser() {
  const old = val("v_old"),
    n = val("v_name"),
    e = (val("v_email") || old).toLowerCase(),
    role = document.getElementById("v_role").value,
    act = document.getElementById("v_active").value === "1",
    p = val("v_pass");
  if (!n || !e) return toast(t("fillAll"));
  if (p && p.length < 8) return toast(t("passShort"));
  await pullHub();
  if (!old && G.users.some((x) => x.email === e)) return toast(t("emailUsed"));
  if (!old && !p) return toast(t("fillAll"));
  const admins = G.users.filter(
    (x) => x.role === "admin" && x.active && x.email !== old,
  ).length;
  if (old && !(role === "admin" && act) && admins === 0)
    return toast(t("lastAdmin"));
  const hash = p ? await sha(e + p) : null;
  const eventIds =
    role === "admin"
      ? []
      : Array.from(document.querySelectorAll("#v_events .chip.on")).map(
          (c) => c.dataset.id,
        );
  await mutateHub(
    () => {
      if (old) {
        const u = G.users.find((x) => x.email === old);
        u.name = n;
        u.role = role;
        u.active = act;
        u.eventIds = eventIds;
        if (hash) u.pass = hash;
      } else
        G.users.push({
          id: uid(),
          name: n,
          email: e,
          role,
          active: act,
          provider: "password",
          pass: hash,
          rec: "",
          created: now(),
          last: "",
          eventIds,
        });
    },
    "actUser",
    `${old ? t("userSaved") : t("userAdded")}: ${n} (${roleLabel(role)})`,
  );
  if (me.email === old) me = G.users.find((x) => x.email === old);
  closeSheet();
  toast(old ? t("userSaved") : t("userAdded"));
  render();
}
async function delUser(email) {
  if (!confirm(t("confirmUser"))) return;
  await pullHub();
  const u = G.users.find((x) => x.email === email);
  if (
    G.users.filter((x) => x.role === "admin" && x.active).length <= 1 &&
    u.role === "admin"
  )
    return toast(t("lastAdmin"));
  await mutateHub(
    () => {
      G.users = G.users.filter((x) => x.email !== email);
    },
    "actUser",
    t("userDel") + ": " + u.name,
  );
  closeSheet();
  toast(t("userDel"));
  render();
}
async function startImp(email) {
  imp = G.users.find((x) => x.email === email);
  await enterResolved(imp, null);
  tab = "dash";
  await mutateHub(() => {}, "actImp", imp.name + " (" + imp.email + ")");
  await saveSession();
  render();
  toast(t("impBanner", { n: imp.name }));
}

/* ================= log aktivitas ================= */
function logRowCardHtml(l) {
  const isDel = /Delete|Hapus/.test(t(l.act)),
    isVer = /Verif/.test(t(l.act));
  return `<div class="drill-row">
    <div class="drill-row-bar" style="background:${isDel ? "var(--red)" : isVer ? "var(--green)" : "var(--blue)"}"></div>
    <div style="flex:1;min-width:0">
      <div class="rowsp" style="gap:8px;align-items:flex-start">
        <div style="min-width:0">
          <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(l.user)}</div>
          <div class="hint" style="margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(l.email)}${l.asAdmin ? " · " + t("byAdmin", { n: esc(l.asAdmin) }) : ""}</div>
        </div>
        <span class="tag ${isDel ? "t-exp" : isVer ? "t-ok" : "t-tic"}" style="flex:none">${t(l.act)}</span>
      </div>
      ${l.det ? `<div class="hint" style="margin-top:4px">${esc(l.det)}</div>` : ""}
      <div class="hint mono" style="margin-top:4px;font-size:12px">${new Date(l.ts).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</div>
    </div>
  </div>`;
}
function logListHtml() {
  const q = logQ.toLowerCase();
  const all = S.logs.filter(
    (l) =>
      !q ||
      [l.user, l.email, t(l.act), l.det].join(" ").toLowerCase().includes(q),
  );
  const { items: rows, page, totalPages } = paginate(all, logPage);
  if (!rows.length) return `<div class="empty">${t("noLog")}</div>`;
  if (narrow()) return rows.map(logRowCardHtml).join("") + pagerHtml(page, totalPages, "setLogPage");
  return `<div style="overflow-x:auto"><table><thead><tr>
    <th>${t("time")}</th><th>${t("user")}</th><th>${t("action")}</th><th>${t("info")}</th></tr></thead><tbody>
    ${rows
      .map(
        (l) => `<tr>
      <td class="hint mono" style="white-space:nowrap">${new Date(l.ts).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</td>
      <td><div style="font-weight:600">${esc(l.user)}</div><div class="hint">${esc(l.email)}${l.asAdmin ? " · " + t("byAdmin", { n: esc(l.asAdmin) }) : ""}</div></td>
      <td><span class="tag ${/Delete|Hapus/.test(t(l.act)) ? "t-exp" : /Verif/.test(t(l.act)) ? "t-ok" : "t-tic"}">${t(l.act)}</span></td>
      <td class="hint">${esc(l.det)}</td></tr>`,
      )
      .join("")}</tbody></table></div>${pagerHtml(page, totalPages, "setLogPage")}`;
}
function vLogs() {
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("logs")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;min-width:220px">
      <input id="logSearchInput" style="max-width:320px" placeholder="${t("searchLog")}" value="${esc(logQ)}"
        oninput="onLogSearch(this.value)">
      <button class="btn ghost sm" onclick="exportLogs()">${t("exportLog")}</button></div></div>
  <div id="logListBody">${logListHtml()}</div></div>`;
}
async function exportLogs() {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(lang === "id" ? "Log" : "Log");
  ws.addRow([t("time"), t("user"), t("email"), t("role"), t("action"), t("info")]);
  styleHeaderRow(ws, 1);
  S.logs.forEach((l) =>
    ws.addRow([new Date(l.ts).toLocaleString(), l.user, l.email, l.role, t(l.act), l.det]),
  );
  ws.columns.forEach((c) => (c.width = 22));
  await downloadWorkbook(wb, `Activity-Log-${today()}.xlsx`);
  toast(t("exportLog"));
}

/* ================= pengaturan (admin) ================= */
function vSet() {
  return `<div class="grid" style="gap:10px">
  <div class="card"><h2 style="font-size:17px;margin-bottom:10px">${t("event")}</h2>
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))">
      <div class="field"><label>${t("evName")}</label><input id="s_ev" value="${esc(D().event)}"></div>
      <div class="field"><label>${t("evDate")}</label><input type="date" id="s_date" value="${D().date || ""}"></div></div>
    <div><button class="btn" onclick="saveSet()">${t("saveSet")}</button></div></div>
  <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(330px,1fr))">
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("catTitle")}</h2>
      <div style="overflow-x:auto"><table style="min-width:640px"><thead><tr><th style="min-width:120px">${t("name")}</th><th style="min-width:110px">${t("groupLbl")}</th><th style="min-width:90px">${t("price")}</th><th style="min-width:70px">${t("qtyLbl")}</th><th style="min-width:80px">${t("quota")}</th><th style="min-width:80px"></th><th></th></tr></thead><tbody>
      ${D()
        .cats.map(
          (
            c,
            i,
          ) => `<tr><td><input value="${esc(c.n)}" onchange="editCat(${i},'n',this.value)"></td>
        <td><select style="width:100%" onchange="editCat(${i},'group',this.value)">
          <option value="income" ${c.group === "income" ? "selected" : ""}>${t("incomeW")}</option>
          <option value="expense" ${c.group === "expense" ? "selected" : ""}>${t("expW")}</option>
        </select></td>
        <td><input type="number" value="${c.p}" onchange="editCat(${i},'p',+this.value)"></td>
        <td style="text-align:center"><input type="checkbox" ${c.hasQty ? "checked" : ""} onchange="editCat(${i},'hasQty',this.checked)"></td>
        <td><input type="number" value="${c.q}" ${c.hasQty ? "" : "disabled"} onchange="editCat(${i},'q',+this.value)"></td>
        <td><button class="btn ghost sm" onclick="openTierManager(${i})">${t("tiersBtn", { n: (c.tiers || []).length })}</button></td>
        <td><button class="btn danger sm" onclick="delCat(${i})">${t("del")}</button></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div><button class="btn ghost sm" style="margin-top:10px" onclick="addCat()">${t("addCat")}</button></div></div>
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("methods")}</h2>
      <div style="overflow-x:auto"><table style="min-width:420px"><thead><tr><th style="min-width:160px">${t("methodName")}</th><th style="min-width:160px">${t("methodType")}</th><th></th></tr></thead><tbody>
      ${D()
        .methods.map(
          (m, i) => `<tr><td><input value="${esc(m.name)}" onchange="editMethod(${i},'name',this.value)"></td>
        <td><select onchange="editMethod(${i},'type',this.value)">
          <option value="bank" ${m.type === "bank" ? "selected" : ""}>${t("methodBank")}</option>
          <option value="cash" ${m.type === "cash" ? "selected" : ""}>${t("methodCash")}</option>
          <option value="cheque" ${m.type === "cheque" ? "selected" : ""}>${t("methodCheque")}</option>
          <option value="debt" ${m.type === "debt" ? "selected" : ""}>${t("methodDebt")}</option>
          <option value="other" ${m.type === "other" ? "selected" : ""}>${t("methodOther")}</option>
        </select></td>
        <td><button class="btn danger sm" onclick="delMethod(${i})">${t("del")}</button></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div><button class="btn ghost sm" style="margin-top:10px" onclick="addMethod()">${t("addMethod")}</button></div></div>
    <div class="card"><h2 style="font-size:17px">${t("tplTitle")}</h2>
      <p class="hint" style="margin:4px 0 8px">${t("tplP")}</p>
      <div style="overflow-x:auto"><table style="min-width:480px"><thead><tr><th style="min-width:150px">${t("field")}</th><th style="min-width:220px">${t("colName")}</th><th></th></tr></thead><tbody>
      ${D()
        .tpl.map(
          (c, i) => `<tr><td style="font-weight:600">${
            c.custom
              ? `<span class="tag t-tic">${t("customCol")}</span>`
              : FL[c.f]
                ? FL[c.f][lang]
                : c.f
          }</td>
        <td><input value="${esc(c.h)}" onchange="editTpl(${i},this.value)"></td>
        <td><button class="btn danger sm" onclick="delTplCol(${i})">${t("del")}</button></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div style="display:flex;gap:8px;margin-top:10px;flex-wrap:wrap">
        <button class="btn ghost sm" onclick="addCustomCol()">${t("addCol")}</button>
        <button class="btn ghost sm" onclick="dlTemplate()">${t("dlTpl")}</button>
        <button class="btn sm" onclick="openImport()">${t("imp")}</button></div></div></div>
  <div class="card"><h2 style="font-size:17px">${t("dataT")}</h2>
    <p class="hint" style="margin:4px 0 10px">${t("dataP")}</p>
    <div><button class="btn danger" onclick="resetAll()">${t("clearAll")}</button></div></div></div>`;
}
async function saveSet() {
  const ev = val("s_ev"),
    dt = val("s_date");
  await mutate(
    () => {
      S.config.event = ev;
      S.config.date = dt;
    },
    "actSet",
    ev,
  );
  await mutateHub(() => {
    const e = G.events.find((x) => x.id === currentEventId);
    if (e) {
      e.name = ev;
      e.date = dt;
    }
  }, null);
  toast(t("setSaved"));
  render();
}
async function addMethod() {
  await mutate(() => {
    S.config.methods.push({ id: uid(), name: "—", type: "bank" });
  }, "actSet", t("addMethod"));
  render();
}
async function editMethod(i, k, v) {
  await mutate(() => {
    S.config.methods[i][k] = v;
  }, "actSet", t("methods") + ": " + S.config.methods[i].name);
}
async function delMethod(i) {
  const n = S.config.methods[i].name;
  await mutate(() => {
    S.config.methods.splice(i, 1);
  }, "actSet", t("del") + ": " + n);
  render();
}
async function editCat(i, k, v) {
  await mutate(
    () => {
      S.config.cats[i][k] = v;
    },
    "actSet",
    t("catTitle") + ": " + S.config.cats[i].n,
  );
}
async function editTpl(i, v) {
  await mutate(
    () => {
      S.config.tpl[i].h = v;
    },
    "actSet",
    t("tplTitle") + ": " + v,
  );
}
async function addCustomCol() {
  const f = "c_" + uid();
  // nama default harus unik dari kolom lain - Excel Table tidak mengizinkan
  // dua kolom dengan nama sama (tabel akan dianggap rusak dan "diperbaiki"
  // Excel dengan cara membuang fitur Table/AutoFilter-nya)
  const base = t("customCol");
  let h = base,
    n = 1;
  while (S.config.tpl.some((c) => c.h === h)) h = `${base} ${++n}`;
  await mutate(
    () => {
      S.config.tpl.push({ f, h, custom: true });
    },
    "actSet",
    t("addCol"),
  );
  render();
}
async function delTplCol(i) {
  const c = S.config.tpl[i];
  await mutate(
    () => {
      S.config.tpl.splice(i, 1);
    },
    "actSet",
    t("del") + ": " + c.h,
  );
  render();
}
async function addCat() {
  await mutate(
    () => {
      S.config.cats.push({ id: uid(), n: "—", group: "income", hasQty: false, p: 0, q: 0 });
    },
    "actSet",
    t("addCat"),
  );
  render();
}
// kategori bertingkat (mis. "Pembelian Kursi" -> Platinum/Gold/Silver, tiap
// tingkat harga sendiri) - dikelola lewat sheet kecil per kategori, bukan
// dijejalkan ke tabel kategori yg sudah padat kolomnya
function openTierManager(catIdx) {
  const c = S.config.cats[catIdx];
  if (!c) return;
  const tieredCats = S.config.cats.filter((tc) => tc.tiers && tc.tiers.length);
  sheet(`<div class="rowsp" style="margin-bottom:10px"><h2 style="font-size:19px">${t("tiersFor", { n: esc(c.n) })}</h2>${closeBtn()}</div>
    <p class="hint" style="margin:0 0 12px">${t("tiersHint")}</p>
    <div style="overflow-x:auto"><table><thead><tr><th>${t("tierName")}</th><th>${t("price")}</th><th></th></tr></thead><tbody>
    ${
      (c.tiers || [])
        .map(
          (tr, i) => `<tr><td><input value="${esc(tr.name)}" onchange="editTier(${catIdx},${i},'name',this.value)"></td>
      <td><input type="number" value="${tr.p}" onchange="editTier(${catIdx},${i},'p',+this.value)"></td>
      <td><button class="btn danger sm" onclick="delTier(${catIdx},${i})">${t("del")}</button></td></tr>`,
        )
        .join("") || `<tr><td colspan="3" class="empty">${t("noneYet")}</td></tr>`
    }
    </tbody></table></div>
    <div style="margin-top:10px"><button class="btn ghost sm" onclick="addTier(${catIdx})">${t("addTier")}</button></div>
    <h2 style="font-size:16px;margin-top:22px">${t("bonusRulesTitle")}</h2>
    <p class="hint" style="margin:4px 0 12px">${t("bonusRulesHint")}</p>
    ${
      tieredCats.length
        ? `<div style="overflow-x:auto"><table><thead><tr><th>${t("bonusMinAmt")}</th><th>${t("bonusTargetCat")}</th><th>${t("bonusTargetTier")}</th><th>${t("bonusQty")}</th><th></th></tr></thead><tbody>
      ${
        (c.bonusRules || [])
          .map((r, i) => {
            const targetCat = tieredCats.find((tc) => tc.id === r.targetCatId) || tieredCats[0];
            return `<tr>
          <td><input type="number" value="${r.minAmount || 0}" onchange="editBonusRule(${catIdx},${i},'minAmount',+this.value)"></td>
          <td><select onchange="editBonusRule(${catIdx},${i},'targetCatId',this.value)">
            ${tieredCats.map((tc) => `<option value="${tc.id}" ${tc.id === (r.targetCatId || targetCat?.id) ? "selected" : ""}>${esc(tc.n)}</option>`).join("")}
          </select></td>
          <td><select onchange="editBonusRule(${catIdx},${i},'targetTier',this.value)">
            ${(targetCat?.tiers || []).map((tr) => `<option ${tr.name === r.targetTier ? "selected" : ""}>${esc(tr.name)}</option>`).join("")}
          </select></td>
          <td><input type="number" value="${r.freeQty || 0}" onchange="editBonusRule(${catIdx},${i},'freeQty',+this.value)"></td>
          <td><button class="btn danger sm" onclick="delBonusRule(${catIdx},${i})">${t("del")}</button></td></tr>`;
          })
          .join("") || `<tr><td colspan="5" class="empty">${t("noneYet")}</td></tr>`
      }
      </tbody></table></div>
      <div style="margin-top:10px"><button class="btn ghost sm" onclick="addBonusRule(${catIdx})">${t("addBonusRule")}</button></div>`
        : `<p class="hint">${t("bonusNeedsTiers")}</p>`
    }
    <div class="rowsp" style="margin-top:16px"><span></span><button class="btn" onclick="closeSheet()">${t("close")}</button></div>`);
}
async function addBonusRule(catIdx) {
  const tieredCats = S.config.cats.filter((tc) => tc.tiers && tc.tiers.length);
  if (!tieredCats.length) return;
  await mutate(
    () => {
      const c = S.config.cats[catIdx];
      if (!c.bonusRules) c.bonusRules = [];
      c.bonusRules.push({
        id: uid(),
        minAmount: 0,
        targetCatId: tieredCats[0].id,
        targetTier: tieredCats[0].tiers[0].name,
        freeQty: 1,
      });
    },
    "actSet",
    t("addBonusRule"),
  );
  closeSheet();
  openTierManager(catIdx);
  render();
}
async function editBonusRule(catIdx, ruleIdx, k, v) {
  await mutate(
    () => {
      const rule = S.config.cats[catIdx].bonusRules[ruleIdx];
      rule[k] = v;
      // ganti target kategori -> tier lama mungkin tak ada di kategori baru,
      // reset ke tier pertamanya
      if (k === "targetCatId") {
        const tc = S.config.cats.find((c) => c.id === v);
        rule.targetTier = tc?.tiers?.[0]?.name || "";
      }
    },
    "actSet",
    t("catTitle"),
  );
  if (k === "targetCatId") {
    closeSheet();
    openTierManager(catIdx);
  }
}
async function delBonusRule(catIdx, ruleIdx) {
  await mutate(
    () => {
      S.config.cats[catIdx].bonusRules.splice(ruleIdx, 1);
    },
    "actSet",
    t("del"),
  );
  closeSheet();
  openTierManager(catIdx);
  render();
}
async function addTier(catIdx) {
  await mutate(
    () => {
      const c = S.config.cats[catIdx];
      if (!c.tiers) c.tiers = [];
      c.tiers.push({ id: uid(), name: "—", p: 0 });
    },
    "actSet",
    t("addTier"),
  );
  closeSheet();
  openTierManager(catIdx);
  render();
}
async function editTier(catIdx, tierIdx, k, v) {
  await mutate(
    () => {
      S.config.cats[catIdx].tiers[tierIdx][k] = v;
    },
    "actSet",
    t("catTitle"),
  );
}
async function delTier(catIdx, tierIdx) {
  await mutate(
    () => {
      S.config.cats[catIdx].tiers.splice(tierIdx, 1);
    },
    "actSet",
    t("del"),
  );
  closeSheet();
  openTierManager(catIdx);
  render();
}
async function delCat(i) {
  const n = S.config.cats[i].n;
  await mutate(
    () => {
      S.config.cats.splice(i, 1);
    },
    "actSet",
    t("del") + ": " + n,
  );
  render();
}
async function resetAll() {
  if (!confirm(t("confirmAll"))) return;
  const n = S.tx.length;
  await mutate(
    () => {
      S.tx = [];
    },
    "actClear",
    n + " " + t("trans"),
  );
  render();
  toast(t("cleared"));
}

/* ================= laporan: periode & styling ================= */
const MONTHS = {
  id: [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
  ],
  en: [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ],
};
function weekOfMonth(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  const first = new Date(d.getFullYear(), d.getMonth(), 1);
  const wd = (first.getDay() + 6) % 7; // Senin=0
  return Math.ceil((d.getDate() + wd) / 7);
}
function periodLabel(dateStr, kind) {
  if (kind === "all") return t("perAllLabel");
  const d = new Date(dateStr + "T00:00:00");
  const mn = MONTHS[lang][d.getMonth()];
  return kind === "month"
    ? `${lang === "id" ? "Periode" : "Period"} ${mn} ${d.getFullYear()}`
    : `${lang === "id" ? "Periode" : "Period"} ${mn} W-${weekOfMonth(dateStr)}`;
}
function txInPeriod(x, kind, refDate) {
  if (kind === "all" || !refDate) return true;
  const d = new Date(x.date + "T00:00:00"),
    r = new Date(refDate + "T00:00:00");
  if (d.getFullYear() !== r.getFullYear() || d.getMonth() !== r.getMonth())
    return false;
  return kind === "month" ? true : weekOfMonth(x.date) === weekOfMonth(refDate);
}
function cellValue(x, c, no) {
  if (c.f === "no") return no;
  if (c.f === "type") return { income: t("incomeW"), expense: t("expW") }[x.type];
  if (c.f === "cat") return catLabel(x);
  if (c.f === "debit") return x.type === "expense" ? null : x.amount;
  if (c.f === "credit") return x.type === "expense" ? x.amount : null;
  if (c.f === "status") return x.status === "verified" ? t("inRek") : t("stW");
  if (c.custom) return (x.custom && x.custom[c.f]) || "";
  return x[c.f] ?? "";
}
function addReportTitle(ws, subtitle) {
  const n = D().tpl.length;
  ws.mergeCells(1, 1, 1, n);
  const c1 = ws.getCell(1, 1);
  c1.value = lang === "id" ? "LAPORAN KEUANGAN" : "FINANCIAL REPORT";
  c1.font = { bold: true, size: 15 };
  c1.alignment = { horizontal: "center" };
  ws.mergeCells(2, 1, 2, n);
  const c2 = ws.getCell(2, 1);
  c2.value = subtitle;
  c2.font = { bold: true, size: 11, color: { argb: "FF6B6B66" } };
  c2.alignment = { horizontal: "center" };
  ws.getRow(1).height = 22;
}
// Excel Table mensyaratkan nama kolom unik - kalau ada dua kolom bernama sama
// (mis. beberapa "Kustom" yang belum diganti nama) file akan dianggap rusak
// oleh Excel dan Table/AutoFilter-nya dibuang otomatis. Ini jaring pengaman
// terakhir supaya export tidak pernah menghasilkan file yang rusak, apa pun
// penyebab duplikatnya.
function uniqueHeaders(cols) {
  const seen = {};
  return cols.map((c) => {
    const base = c.h || c.f;
    seen[base] = (seen[base] || 0) + 1;
    return seen[base] === 1 ? base : `${base} (${seen[base]})`;
  });
}
function addTxTable(ws, txList, tableName) {
  const cols = D().tpl;
  const headers = uniqueHeaders(cols);
  // Excel Table tidak boleh punya 0 baris data (UI-nya sendiri tidak
  // mengizinkan itu) - kalau periode yang dipilih kebetulan tidak ada
  // transaksinya, ws.addTable() menghasilkan definisi tabel yang tidak valid
  // dan seluruh file dianggap rusak oleh Excel. Kalau kosong, tulis header
  // biasa saja (tanpa fitur Table/AutoFilter) supaya filenya tetap valid.
  if (txList.length) {
    ws.addTable({
      name: tableName + Math.floor(Math.random() * 1e6),
      ref: "A4",
      headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: cols.map((c, i) => ({ name: headers[i], filterButton: true })),
      rows: txList.map((x, i) => cols.map((c) => cellValue(x, c, i + 1))),
    });
  } else {
    const row = ws.getRow(4);
    headers.forEach((h, i) => (row.getCell(i + 1).value = h));
    styleHeaderRow(ws, 4);
  }
  cols.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    col.width = Math.max(c.h.length + 4, c.f === "note" ? 24 : c.f === "name" ? 20 : 13);
    if (c.f === "debit" || c.f === "credit") col.numFmt = "#,##0";
  });
}
function styleHeaderRow(ws, rowIdx) {
  const row = ws.getRow(rowIdx);
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
    cell.alignment = { vertical: "middle" };
  });
}
async function downloadWorkbook(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ================= impor ================= */
// kolom yg wajib diisi saat impor (dicek berdasar field key, bukan teks
// header, supaya tetap benar walau admin sudah mengubah nama kolomnya)
const MANDATORY_FIELDS = ["date", "name", "cat", "amount", "bank", "status"];
async function dlTemplate() {
  const id = lang === "id";
  const cats = D().cats;
  const cols = D().tpl;
  const headers = uniqueHeaders(cols).map((h, i) =>
    MANDATORY_FIELDS.includes(cols[i].f) && !h.endsWith("*") ? h + "*" : h,
  );
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Template");
  ws.addRow(headers);
  styleHeaderRow(ws, 1);

  const colLetter = (i) => String.fromCharCode(65 + i);
  const dateIdx = cols.findIndex((c) => c.f === "date");
  const catIdx = cols.findIndex((c) => c.f === "cat");
  const noteIdx = cols.findIndex((c) => c.f === "note");
  const note2Idx = cols.findIndex((c) => c.f === "note2");
  const amountIdx = cols.findIndex((c) => c.f === "amount");
  const statusIdx = cols.findIndex((c) => c.f === "status");
  const bankIdx = cols.findIndex((c) => c.f === "bank");
  const bankNameIdx = cols.findIndex((c) => c.f === "bankName");
  const phoneIdx = cols.findIndex((c) => c.f === "phone");
  const tieredCats = cats.filter((c) => c.tiers?.length);

  const exCat = tieredCats[0] || cats.find((c) => c.group === "income" && !c.hasQty) || cats[0];
  const exMethod = D().methods.find((m) => m.type === "bank") || D().methods[0];
  const exampleNote = id ? "iuran bulan Juli" : "July dues";
  const example = {
    name: "Michael",
    phone: "0812xxxxxxxx",
    cat: exCat?.n || "",
    note: exCat?.tiers?.length ? exCat.tiers[0].name : exampleNote,
    note2: exCat?.tiers?.length ? 1 : exampleNote,
    bank: exMethod?.name || "",
    bankName: exMethod?.type === "bank" ? "BCA" : "",
    status: t("stV"),
  };
  const row2 = ws.addRow(cols.map((c) => (c.custom || c.f === "date" || c.f === "amount" ? "" : (example[c.f] ?? ""))));
  if (dateIdx >= 0) row2.getCell(dateIdx + 1).value = new Date(`${today()}T00:00:00`);
  if (amountIdx >= 0) row2.getCell(amountIdx + 1).value = 1000000;

  // tanggal HARUS dd/mm/yyyy: format tampilan kolom dipaksa dd/mm/yyyy +
  // validasi tipe tanggal (menolak teks non-tanggal), bukan cuma contoh baris
  if (dateIdx >= 0) {
    const col = ws.getColumn(dateIdx + 1);
    col.numFmt = "dd/mm/yyyy";
    ws.dataValidations.add(`${colLetter(dateIdx)}2:${colLetter(dateIdx)}1000`, {
      type: "date",
      operator: "between",
      formulae: [new Date(2000, 0, 1), new Date(2100, 11, 31)],
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: id ? "Format tanggal salah" : "Invalid date",
      error: id
        ? "Isi tanggal dgn format dd/mm/yyyy (gunakan date picker Excel)."
        : "Enter a valid date in dd/mm/yyyy format (use Excel's date picker).",
    });
  }
  // nominal ditampilkan sbg Rupiah begitu diisi (contoh: 1000000 -> Rp 1.000.000)
  if (amountIdx >= 0) ws.getColumn(amountIdx + 1).numFmt = '"Rp "#,##0';
  // no. whatsapp disimpan sbg teks supaya angka 0 di depan & digit panjang
  // tidak diubah Excel jadi notasi ilmiah / kehilangan leading zero
  if (phoneIdx >= 0) ws.getColumn(phoneIdx + 1).numFmt = "@";

  // dropdown Excel utk kategori/status/metode pembayaran/bank, supaya user
  // tinggal pilih drpd salah ketik (kategori yg dipilih menentukan pemasukan
  // vs pengeluaran saat diimpor, lihat parseRows())
  if (catIdx >= 0 && cats.length) {
    const ref = wb.addWorksheet("_ref");
    ref.state = "veryHidden";
    cats.forEach((c, i) => (ref.getCell(i + 1, 1).value = c.n));
    INDONESIA_BANKS.forEach((b, i) => (ref.getCell(i + 1, 2).value = b));
    const tierNames = [...new Set(tieredCats.flatMap((c) => c.tiers.map((tr) => tr.name)))];
    tierNames.forEach((n, i) => (ref.getCell(i + 1, 3).value = n));
    ws.dataValidations.add(`${colLetter(catIdx)}2:${colLetter(catIdx)}1000`, {
      type: "list",
      allowBlank: false,
      formulae: [`_ref!$A$1:$A$${cats.length}`],
      showErrorMessage: true,
      errorTitle: id ? "Kategori tidak dikenali" : "Unknown category",
      error: id ? "Pilih kategori dari daftar dropdown." : "Choose a category from the dropdown list.",
    });
    if (bankNameIdx >= 0) {
      ws.dataValidations.add(`${colLetter(bankNameIdx)}2:${colLetter(bankNameIdx)}1000`, {
        type: "list",
        allowBlank: true,
        formulae: [`_ref!$B$1:$B$${INDONESIA_BANKS.length}`],
      });
      const note = ws.getCell(1, bankNameIdx + 1);
      note.note = id
        ? "Wajib diisi jika Metode Pembayaran = Transfer Bank."
        : "Required only when Payment Method = Transfer Bank.";
    }
    // dropdown tingkat (Platinum/Gold/Silver) di kolom Keterangan hanya utk
    // kategori bertingkat (mis. Pembelian Kursi) - bukan validasi keras (spy
    // kategori lain masih bisa isi catatan bebas spt biasa), cuma kemudahan
    if (tierNames.length && noteIdx >= 0) {
      ws.dataValidations.add(`${colLetter(noteIdx)}2:${colLetter(noteIdx)}1000`, {
        type: "list",
        allowBlank: true,
        showErrorMessage: false,
        formulae: [`_ref!$C$1:$C$${tierNames.length}`],
      });
      const note = ws.getCell(1, noteIdx + 1);
      note.note = id
        ? `Untuk kategori bertingkat (${tieredCats.map((c) => c.n).join(", ")}): isi nama tingkat (${tierNames.join("/")}). Kategori lain: catatan bebas.`
        : `For tiered categories (${tieredCats.map((c) => c.n).join(", ")}): enter the tier name (${tierNames.join("/")}). Other categories: free text.`;
      if (note2Idx >= 0) {
        const note2Cell = ws.getCell(1, note2Idx + 1);
        note2Cell.note = id
          ? `Untuk kategori bertingkat: isi jumlah kursi yang dibeli (angka). Kategori lain: catatan tambahan bebas.`
          : `For tiered categories: enter the number of seats purchased. Other categories: free additional note.`;
      }
    }
  }
  if (statusIdx >= 0) {
    ws.dataValidations.add(`${colLetter(statusIdx)}2:${colLetter(statusIdx)}1000`, {
      type: "list",
      allowBlank: false,
      formulae: [`"${t("stV")},${t("stW")}"`],
    });
  }
  if (bankIdx >= 0 && D().methods.length) {
    ws.dataValidations.add(`${colLetter(bankIdx)}2:${colLetter(bankIdx)}1000`, {
      type: "list",
      allowBlank: false,
      formulae: [`"${D().methods.map((m) => m.name).join(",")}"`],
    });
  }

  cols.forEach((c, i) => {
    ws.getColumn(i + 1).width = Math.max(
      (headers[i] || "").length + 3,
      c.f === "note" || c.f === "note2" ? 22 : c.f === "name" ? 18 : 14,
    );
  });
  await downloadWorkbook(wb, "Template-Import.xlsx");
  toast(t("dlTpl"));
}
function openImport() {
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:20px">${t("imp")}</h2>
    ${closeBtn()}</div>
    <div class="drop" id="drop"><div style="font-size:17px;font-weight:600">${t("impDrop")}</div>
      <div class="hint" style="margin-top:4px">${t("impSub")} · .xlsx .xls .csv</div></div>
    <input type="file" id="file" accept=".xlsx,.xls,.csv" style="display:none">
    <div id="impBody"></div>
    <div style="margin-top:12px"><button class="btn ghost sm" onclick="dlTemplate()">${t("dlTpl")}</button></div>`);
  const d = document.getElementById("drop"),
    f = document.getElementById("file");
  d.onclick = () => f.click();
  d.ondragover = (e) => {
    e.preventDefault();
    d.classList.add("over");
  };
  d.ondragleave = () => d.classList.remove("over");
  d.ondrop = (e) => {
    e.preventDefault();
    d.classList.remove("over");
    readFile(e.dataTransfer.files[0]);
  };
  f.onchange = () => readFile(f.files[0]);
}
function readFile(file) {
  if (!file) return;
  const r = new FileReader();
  r.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
    parseRows(
      XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        header: 1,
        raw: false,
        defval: "",
      }),
    );
  };
  r.readAsArrayBuffer(file);
}
// nama bulan Indonesia & Inggris (+ singkatan umum) - dipakai normDate() utk
// mengenali tanggal spt "25 Juli 2026" yg tidak dikenali Date() bawaan JS
const MONTH_NAMES = {
  jan: 0, januari: 0, january: 0,
  feb: 1, februari: 1, february: 1,
  mar: 2, maret: 2, march: 2,
  apr: 3, april: 3,
  mei: 4, may: 4,
  jun: 5, juni: 5, june: 5,
  jul: 6, juli: 6, july: 6,
  agu: 7, agt: 7, agustus: 7, aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  okt: 9, oktober: 9, oct: 9, october: 9,
  nov: 10, november: 10,
  des: 11, desember: 11, dec: 11, december: 11,
};
function normDate(v) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  // "25 Juli 2026" / "25 July 2026" (juga terima "Juli 25, 2026")
  m = s.match(/^(\d{1,2})\s+([a-zA-Z]+)\.?,?\s+(\d{4})$/);
  if (!m) {
    const m2 = s.match(/^([a-zA-Z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/);
    if (m2) m = [m2[0], m2[2], m2[1], m2[3]];
  }
  if (m) {
    const mo = MONTH_NAMES[m[2].toLowerCase()];
    if (mo !== undefined) {
      return `${m[3]}-${String(mo + 1).padStart(2, "0")}-${String(+m[1]).padStart(2, "0")}`;
    }
  }
  if (/^\d+$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + +s * 864e5);
    return d.toISOString().slice(0, 10);
  }
  const d = new Date(s);
  return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}
// dipakai cuma sbg fallback pas kategori di baris impor tidak dikenali -
// jalur utama sekarang mencocokkan nama kategori ke config.cats (lihat parseRows)
function normType(v) {
  const s = String(v || "").toLowerCase();
  if (/exp|keluar|biaya|beban|pengeluaran/.test(s)) return "expense";
  return "income";
}
function normNum(v) {
  return +String(v ?? "").replace(/[^\d.-]/g, "") || 0;
}
// file yang di-download dari app sendiri (template maupun laporan) punya 3
// baris judul (LAPORAN KEUANGAN / periode / kosong) sebelum baris header -
// cari baris mana yang paling cocok dengan nama kolom di D().tpl, jangan
// asumsikan header selalu di baris pertama (juga tetap benar untuk file
// polos yang headernya memang di baris pertama)
// header di file yg diunduh dari app sendiri punya sufiks "*" utk kolom
// wajib (lihat dlTemplate()) - abaikan sufiks itu saat mencocokkan header,
// supaya cocok terlepas dari apakah tpl tersimpan sudah punya "*" atau belum
const normHeader = (h) => String(h).trim().toLowerCase().replace(/\*\s*$/, "");
function findHeaderRowIndex(rows) {
  const tplHeaders = D().tpl.map((c) => normHeader(c.h));
  let bestIdx = 0,
    bestScore = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = (rows[i] || []).map((h) => normHeader(h));
    const score = row.filter((h) => tplHeaders.includes(h)).length;
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestIdx;
}
function parseRows(rows) {
  const body = document.getElementById("impBody");
  if (!rows || rows.length < 2)
    return (body.innerHTML = `<div class="empty">${t("noCol")}</div>`);
  const headerIdx = findHeaderRowIndex(rows);
  const head = (rows[headerIdx] || []).map((h) => normHeader(h)),
    idx = {};
  D().tpl.forEach((c) => {
    const i = head.indexOf(normHeader(c.h));
    if (i >= 0) idx[c.f] = i;
  });
  if (
    idx.name === undefined &&
    idx.amount === undefined &&
    idx.debit === undefined &&
    idx.credit === undefined
  )
    return (body.innerHTML = `<div class="empty">${t("noCol")}</div>`);
  const get = (r, f) => (idx[f] === undefined ? "" : r[idx[f]]);
  const customCols = D().tpl.filter((c) => c.custom);
  pendingImp = [];
  let bad = 0;
  rows.slice(headerIdx + 1).forEach((r) => {
    if (!r || !r.length || r.every((c) => String(c).trim() === "")) return;
    const nameRaw = String(get(r, "name") || "").trim();
    const rawCat = String(get(r, "cat") || "").trim();
    const bankRaw = String(get(r, "bank") || "").trim();
    const bankNameRaw = String(get(r, "bankName") || "").trim();
    const statusRaw = String(get(r, "status") || "").trim();
    const debit = normNum(get(r, "debit")),
      credit = normNum(get(r, "credit")),
      legacyAmt = normNum(get(r, "amount")),
      amt = debit || credit || legacyAmt,
      dt = normDate(get(r, "date"));
    // utamakan grup kategori yg sudah dikenal (config.cats) drpd teks kolom
    // "type" - sumber kebenaran skrg ada di kategori, bukan lagi ticket/
    // donation/expense; teks "type" cuma fallback kalau kategorinya tak dikenal
    const matchedCat = D().cats.find((c) => c.n.toLowerCase() === rawCat.toLowerCase());
    const matchedMethod = D().methods.find((m) => m.name.toLowerCase() === bankRaw.toLowerCase());
    // kolom Bank cuma wajib kalau metode pembayarannya "bank" (transfer) -
    // utk Tunai/Utang/Cek kosong itu wajar, jadi tidak dianggap baris rusak
    const bankNameMissing = matchedMethod?.type === "bank" && !bankNameRaw;
    // kolom wajib (lihat MANDATORY_FIELDS): tanggal, nama, kategori (harus
    // cocok dgn kategori yg ada), nominal, metode pembayaran (harus cocok),
    // status - baris yg kosong/tak dikenali di salah satunya dilewati
    if (!dt || !nameRaw || !matchedCat || !amt || !matchedMethod || !statusRaw || bankNameMissing) {
      bad++;
      return;
    }
    const ty = matchedCat.group;
    const st = /masuk|verif|lunas|received|paid|ok/i.test(statusRaw) ? "verified" : "pending";
    const matchedBankName = matchedMethod.type === "bank"
      ? INDONESIA_BANKS.find((b) => b.toLowerCase() === bankNameRaw.toLowerCase()) || bankNameRaw
      : "";
    // kategori bertingkat (mis. Pembelian Kursi): kolom Keterangan berisi nama
    // tingkat & Keterangan Tambahan berisi jumlah kursi, bukan catatan bebas -
    // lihat hint di dlTemplate()
    const noteRaw = String(get(r, "note") || "");
    const note2Raw = String(get(r, "note2") || "");
    let tier = "", seats = 0, note = noteRaw, note2 = note2Raw;
    if (matchedCat.tiers?.length) {
      const matchedTier = matchedCat.tiers.find((tr) => tr.name.toLowerCase() === noteRaw.trim().toLowerCase());
      tier = matchedTier?.name || "";
      seats = normNum(note2Raw);
      note = "";
      note2 = "";
    } else if (matchedCat.hasQty) {
      // kategori hasQty TANPA tingkat (mis. tpl lama dgn kolom "Kursi"
      // terpisah) - jumlah kursinya tetap dibaca dari kolom generik "seats",
      // bukan dari Keterangan Tambahan (yg cuma dipakai utk kategori bertingkat)
      seats = normNum(get(r, "seats"));
    }
    const custom = {};
    customCols.forEach((c) => {
      custom[c.f] = String(get(r, c.f) || "");
    });
    pendingImp.push({
      id: uid(),
      type: ty,
      date: dt,
      name: nameRaw || "—",
      phone: String(get(r, "phone") || ""),
      cat: matchedCat.n,
      tier,
      seats,
      amount: amt,
      bank: matchedMethod.name,
      bankName: matchedBankName,
      methodType: matchedMethod.type,
      chequeNo: "",
      chequeBank: "",
      chequeDate: "",
      note,
      note2,
      proof: "",
      status: st,
      by: acting().name,
      custom,
    });
  });
  const lbl = { income: t("incomeW"), expense: t("expW") };
  body.innerHTML = `<div style="margin-top:12px"><div class="rowsp"><span class="label">${t("prev")}</span>
    <span class="hint">${t("rowsOk", { n: pendingImp.length })}${bad ? " · " + t("rowsBad", { n: bad }) : ""}</span></div>
    <div class="scroll" style="max-height:230px;border:1px solid var(--line);border-radius:10px;margin-top:6px;padding:0 10px">
    <table><thead><tr><th>${t("date")}</th><th>${t("name")}</th><th>${t("detail")}</th><th style="text-align:right">${t("amount")}</th></tr></thead><tbody>
    ${pendingImp
      .map(
        (
          x,
        ) => `<tr><td class="mono hint">${x.date.slice(5)}</td><td style="font-weight:600">${esc(x.name)}</td>
      <td>${lbl[x.type]} ${esc(catLabel(x))} ${x.seats ? "· " + x.seats : ""}</td>
      <td class="mono" style="text-align:right">${rp(x.amount)}</td></tr>`,
      )
      .join("")}</tbody></table></div>
    <button class="btn wide" style="margin-top:12px" onclick="commitImport()">${t("doImp", { n: pendingImp.length })}</button></div>`;
}
async function commitImport() {
  const n = pendingImp.length,
    rows = pendingImp.slice();
  pendingImp = [];
  await mutate(
    () => {
      S.tx = S.tx.concat(rows);
    },
    "actImport",
    n + " " + t("trans"),
  );
  closeSheet();
  toast(t("imported", { n }));
  tab = "tx";
  render();
}

/* ================= excel ================= */
function openExportPeriod() {
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${t("xls")}</h2>${closeBtn()}</div>
    <div class="seg" id="expSeg" style="margin-bottom:14px">${[
      ["week", t("perWeek")],
      ["month", t("perMonth")],
      ["all", t("perAll")],
    ]
      .map(
        ([k, l]) =>
          `<button class="${exportKind === k ? "on" : ""}" data-k="${k}" onclick="setExportKind('${k}')">${l}</button>`,
      )
      .join("")}</div>
    <div class="field" id="expDateWrap" style="${exportKind === "all" ? "display:none" : ""}">
      <label>${t("perDate")}</label>
      <input type="date" id="exp_date" value="${today()}" oninput="updateExportPreview()"></div>
    <div class="hint mono" id="expPreview" style="margin-bottom:16px">${periodLabel(today(), exportKind)}</div>
    <button class="btn wide" onclick="doExportExcel()">${t("xls")}</button>`);
}
function setExportKind(k) {
  exportKind = k;
  document
    .querySelectorAll("#expSeg button")
    .forEach((b) => b.classList.toggle("on", b.dataset.k === k));
  document.getElementById("expDateWrap").style.display =
    k === "all" ? "none" : "block";
  updateExportPreview();
}
function updateExportPreview() {
  const d = document.getElementById("exp_date")?.value || today();
  document.getElementById("expPreview").textContent = periodLabel(d, exportKind);
}
async function doExportExcel() {
  const refDate = document.getElementById("exp_date")?.value || today();
  const kind = exportKind;
  closeSheet();
  await exportExcel(kind, refDate);
}
// laporan-laporan di bawah ini dibuat spt neraca sederhana (bagian
// berjudul, item diindentasi, baris total bergaris atas) - bukan tabel
// transaksi mentah. Urutan sheet: (1) keseluruhan, (2) penerimaan, (3)
// pengeluaran, (4+) satu sheet per kategori, terakhir tabel transaksi mentah
// sbg referensi tambahan.
function moneyCell(ws, row, col, value) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.numFmt = "#,##0";
  return cell;
}
function statementHeader(ws, subtitle, spanCols = 3) {
  ws.mergeCells(1, 1, 1, spanCols);
  const c1 = ws.getCell(1, 1);
  c1.value = lang === "id" ? "LAPORAN KEUANGAN" : "FINANCIAL STATEMENT";
  c1.font = { bold: true, size: 15 };
  c1.alignment = { horizontal: "center" };
  ws.mergeCells(2, 1, 2, spanCols);
  const c2 = ws.getCell(2, 1);
  c2.value = D().event;
  c2.font = { bold: true, size: 12 };
  c2.alignment = { horizontal: "center" };
  ws.mergeCells(3, 1, 3, spanCols);
  const c3 = ws.getCell(3, 1);
  c3.value = subtitle;
  c3.font = { size: 10, color: { argb: "FF6B6B66" } };
  c3.alignment = { horizontal: "center" };
  ws.getRow(1).height = 22;
  return 5;
}
const topBorder = (ws, row, style = "thin") =>
  ws.getRow(row).eachCell({ includeEmpty: true }, (cell) => (cell.border = { top: { style } }));
function catAmount(c, kind, refDate) {
  return S.tx
    .filter((x) => x.status === "verified" && x.cat === c.n && txInPeriod(x, kind, refDate))
    .reduce((a, b) => a + b.amount, 0);
}
function uniqueSheetName(base, usedNames) {
  const clean = String(base).replace(/[\\/?*[\]:]/g, "-").slice(0, 28) || "Sheet";
  let name = clean,
    n = 1;
  while (usedNames.has(name.toLowerCase())) name = `${clean.slice(0, 24)} (${++n})`;
  usedNames.add(name.toLowerCase());
  return name;
}
function buildOverallSheet(wb, kind, refDate, usedNames) {
  const id = lang === "id";
  const ws = wb.addWorksheet(uniqueSheetName(id ? "Laporan Keseluruhan" : "Overall Report", usedNames));
  let row = statementHeader(ws, periodLabel(refDate, kind));
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value =
    (id ? "Dicetak" : "Generated") + ": " + new Date().toLocaleString() + " · " + (id ? "Oleh" : "By") + ": " + acting().name;
  ws.getCell(row, 1).font = { size: 9, italic: true, color: { argb: "FF9A9A93" } };
  row += 2;

  const incomeCats = D().cats.filter((c) => c.group === "income");
  const expenseCats = D().cats.filter((c) => c.group === "expense");

  ws.getCell(row, 1).value = t("incomeW").toUpperCase();
  ws.getCell(row, 1).font = { bold: true, size: 12 };
  row++;
  let incomeTotal = 0;
  incomeCats.forEach((c) => {
    const v = catAmount(c, kind, refDate);
    incomeTotal += v;
    ws.getCell(row, 1).value = "    " + c.n;
    moneyCell(ws, row, 3, v);
    row++;
  });
  ws.getCell(row, 1).value = id ? "Total Penerimaan" : "Total Income";
  ws.getCell(row, 1).font = { bold: true };
  moneyCell(ws, row, 3, incomeTotal).font = { bold: true };
  topBorder(ws, row);
  row += 2;

  ws.getCell(row, 1).value = t("exp").toUpperCase();
  ws.getCell(row, 1).font = { bold: true, size: 12 };
  row++;
  let expenseTotal = 0;
  expenseCats.forEach((c) => {
    const v = catAmount(c, kind, refDate);
    expenseTotal += v;
    ws.getCell(row, 1).value = "    " + c.n;
    moneyCell(ws, row, 3, v);
    row++;
  });
  ws.getCell(row, 1).value = id ? "Total Pengeluaran" : "Total Expense";
  ws.getCell(row, 1).font = { bold: true };
  moneyCell(ws, row, 3, expenseTotal).font = { bold: true };
  topBorder(ws, row);
  row += 2;

  ws.getCell(row, 1).value = t("net").toUpperCase();
  ws.getCell(row, 1).font = { bold: true, size: 13 };
  moneyCell(ws, row, 3, incomeTotal - expenseTotal).font = { bold: true, size: 13 };
  topBorder(ws, row, "double");

  ws.getColumn(1).width = 36;
  ws.getColumn(2).width = 4;
  ws.getColumn(3).width = 20;
}
function buildGroupSheet(wb, group, kind, refDate, usedNames) {
  const id = lang === "id";
  const label = group === "income" ? t("incomeW") : t("exp");
  const ws = wb.addWorksheet(uniqueSheetName(label, usedNames));
  let row = statementHeader(ws, label + " · " + periodLabel(refDate, kind));
  const cats = D().cats.filter((c) => c.group === group);
  ws.getCell(row, 1).value = t("txCat");
  ws.getCell(row, 2).value = id ? "Jml. Transaksi" : "Transactions";
  ws.getCell(row, 3).value = id ? "Total" : "Total";
  styleHeaderRow(ws, row);
  row++;
  let grand = 0,
    grandN = 0;
  cats.forEach((c) => {
    const tx = S.tx.filter((x) => x.status === "verified" && x.cat === c.n && txInPeriod(x, kind, refDate));
    const v = tx.reduce((a, b) => a + b.amount, 0);
    grand += v;
    grandN += tx.length;
    ws.getCell(row, 1).value = c.n;
    ws.getCell(row, 2).value = tx.length;
    moneyCell(ws, row, 3, v);
    row++;
  });
  ws.getCell(row, 1).value = id ? `Total ${label}` : `Total ${label}`;
  ws.getCell(row, 1).font = { bold: true };
  ws.getCell(row, 2).value = grandN;
  ws.getCell(row, 2).font = { bold: true };
  moneyCell(ws, row, 3, grand).font = { bold: true };
  topBorder(ws, row);
  ws.getColumn(1).width = 32;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 18;
}
function buildCategorySheet(wb, cat, usedNames, kind, refDate) {
  const id = lang === "id";
  const ws = wb.addWorksheet(uniqueSheetName(cat.n, usedNames));
  let row = statementHeader(ws, cat.n + " · " + periodLabel(refDate, kind), 6);
  const tx = S.tx
    .filter((x) => x.status === "verified" && x.cat === cat.n && txInPeriod(x, kind, refDate))
    .sort((a, b) => a.date.localeCompare(b.date));
  [t("date"), t("name"), t("note"), t("paymentMethod"), t("payStatus"), t("amount")].forEach(
    (h, i) => (ws.getCell(row, i + 1).value = h),
  );
  styleHeaderRow(ws, row);
  row++;
  let total = 0;
  tx.forEach((x) => {
    ws.getCell(row, 1).value = x.date;
    ws.getCell(row, 2).value = x.name;
    ws.getCell(row, 3).value = x.tier ? `${x.note || ""} (${x.tier})`.trim() : x.note || "";
    ws.getCell(row, 4).value = x.bank + (x.bankName ? " - " + x.bankName : "");
    ws.getCell(row, 5).value = t("inRek");
    moneyCell(ws, row, 6, x.amount);
    total += x.amount;
    row++;
  });
  if (!tx.length) {
    ws.getCell(row, 1).value = t("noneYet");
    ws.getCell(row, 1).font = { italic: true, color: { argb: "FF9A9A93" } };
    row++;
  }
  ws.getCell(row, 5).value = id ? "Total" : "Total";
  ws.getCell(row, 5).font = { bold: true };
  moneyCell(ws, row, 6, total).font = { bold: true };
  topBorder(ws, row);
  const widths = [14, 20, 26, 20, 12, 16];
  ws.columns.forEach((c, i) => (c.width = widths[i] || 16));
}
async function exportExcel(kind, refDate) {
  const id = lang === "id";
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = "Kas Acara";
  const usedNames = new Set();
  buildOverallSheet(wb, kind, refDate, usedNames);
  buildGroupSheet(wb, "income", kind, refDate, usedNames);
  buildGroupSheet(wb, "expense", kind, refDate, usedNames);
  D().cats.forEach((c) => buildCategorySheet(wb, c, usedNames, kind, refDate));
  const txRows = S.tx
    .filter((x) => txInPeriod(x, kind, refDate))
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date));
  const wsTx = wb.addWorksheet(uniqueSheetName(id ? "Transaksi" : "Transactions", usedNames));
  addReportTitle(wsTx, periodLabel(refDate, kind));
  addTxTable(wsTx, txRows, "TabelTransaksi");
  await downloadWorkbook(
    wb,
    `${id ? "Laporan" : "Report"}-${(D().event || "Acara").replace(/\s+/g, "-")}-${refDate || today()}.xlsx`,
  );
  await mutate(() => {}, "actExport", D().event + " · " + rp(sums().net));
  toast(t("xls"));
}
Object.assign(window, {
  createFirst,
  doSignIn,
  doSignUp,
  doGoogle,
  doForgot,
  doSignOut,
  googleFlow,
  setLang,
  toggleTheme,
  go,
  setAuthMode,
  setPanel,
  setFilter,
  setAdminTab,
  onLogSearch,
  setLogPage,
  setTxSort,
  onTxSearch,
  setTxPage,
  setBoardBuyPage,
  setBoardDonPage,
  setHubEventsPage,
  onHubEventsSearch,
  setHubEventsSort,
  setHubEventsFilter,
  setHubStaffPage,
  onHubStaffSearch,
  setHubStaffSort,
  setHubStaffFilter,
  openTx,
  saveTx,
  delTx,
  verify,
  onCatChange,
  setTxTab,
  onTierChange,
  openTierManager,
  addTier,
  editTier,
  delTier,
  addBonusRule,
  editBonusRule,
  delBonusRule,
  updateBonusHint,
  recalc,
  stepSeats,
  prev,
  onProofChange,
  viewProof,
  viewProofPreview,
  viewProofById,
  closeSheet,
  openImport,
  dlTemplate,
  commitImport,
  openMe,
  stopImp,
  openUser,
  saveUser,
  delUser,
  startImp,
  toggleEventChip,
  switchEvent,
  goHub,
  setHubTab,
  openNewEvent,
  createEventSubmit,
  openDuplicateEvent,
  submitDuplicateEvent,
  toggleArchive,
  loadMonitor,
  goDashEditor,
  openWidgetForm,
  onWidgetTypeChange,
  onWidgetGroupChange,
  drilldownKpi,
  drilldownCat,
  drilldownPerson,
  drilldownDate,
  onDrillSearch,
  setDrillPage,
  submitWidgetForm,
  deleteDashWidget,
  saveDashLayout,
  saveSet,
  editCat,
  editTpl,
  addCustomCol,
  delTplCol,
  addCat,
  delCat,
  addMethod,
  editMethod,
  delMethod,
  onMethodChange,
  resetAll,
  openExportPeriod,
  setExportKind,
  updateExportPreview,
  doExportExcel,
  exportLogs,
  toggleEye,
});
boot();
