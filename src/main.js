import "./style.css";
import "./storage.js";
import * as XLSX from "xlsx";
import Chart from "chart.js/auto";

let _ExcelJS = null;
async function loadExcelJS() {
  if (!_ExcelJS) _ExcelJS = (await import("exceljs")).default;
  return _ExcelJS;
}

/* ================= penyimpanan ================= */
// prefix isolasi untuk verifikasi/testing, jangan pernah sentuh key produksi
const KEY_PREFIX = import.meta.env.VITE_KEY_PREFIX || "";
const SKEY = KEY_PREFIX + "ct_shared_v4",
  HKEY = KEY_PREFIX + "ct_events_index_v1",
  LKEY = KEY_PREFIX + "ct_session_v4";
const eventKey = (id) => KEY_PREFIX + "ct_event_" + id + "_v1";
const DEFAULT_TPL = [
  { f: "no", h: "No" },
  { f: "name", h: "Nama" },
  { f: "phone", h: "No. WhatsApp" },
  { f: "type", h: "Jenis" },
  { f: "cat", h: "Kategori" },
  { f: "seats", h: "Jumlah Kursi" },
  { f: "debit", h: "Debit (Rp)" },
  { f: "credit", h: "Kredit (Rp)" },
  { f: "bank", h: "Rekening" },
  { f: "status", h: "Status" },
  { f: "note", h: "Catatan" },
  { f: "date", h: "Tanggal" },
];
const CASH_DENOMS = [100000, 50000, 20000, 10000, 5000, 2000, 1000];
// id statis (bukan uid()) supaya bisa dipanggil saat modul dimuat, sebelum
// uid() sendiri didefinisikan lebih bawah di file ini
const DEFAULT_METHODS = () => [
  { id: "m_bca", name: "BCA", type: "bank" },
  { id: "m_livin", name: "Livin Mandiri", type: "bank" },
  { id: "m_bri", name: "BRI", type: "bank" },
  { id: "m_tunai", name: "Tunai", type: "cash" },
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
      { n: "Reguler", p: 150000, q: 300 },
      { n: "VIP", p: 400000, q: 80 },
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
let chart1 = null,
  pendingImp = [],
  logQ = "",
  txQ = "",
  txSort = { k: "date", dir: "desc" },
  exportKind = "week",
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
const today = () => new Date().toISOString().slice(0, 10);
const narrow = () => window.innerWidth < 960;
const now = () => new Date().toISOString();
const acting = () => imp || me; // identitas yang sedang dipakai
const isAdmin = () => me && me.role === "admin" && !imp; // hak admin hilang saat menyamar
// potong array jadi satu halaman (dipakai semua tampilan list: transaksi,
// log, peringkat, acara, staff) supaya tidak perlu scroll ratusan baris
function paginate(arr, page, size = PAGE_SIZE) {
  const totalPages = Math.max(1, Math.ceil(arr.length / size));
  const p = Math.min(Math.max(1, page), totalPages);
  return { items: arr.slice((p - 1) * size, p * size), page: p, totalPages, total: arr.length };
}
function pagerHtml(page, totalPages, setter) {
  if (totalPages <= 1) return "";
  return `<div class="rowsp" style="margin-top:10px;flex:none;justify-content:center;gap:14px">
    <button class="btn ghost sm icon" ${page <= 1 ? "disabled" : ""} onclick="${setter}(${page - 1})" aria-label="prev">‹</button>
    <span class="hint mono">${page} / ${totalPages}</span>
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
      config: { event: name, date: date || "", cats: [], methods: [], tpl: DEFAULT_TPL },
      tx: [],
      logs: [],
    }),
    true,
  );
  return { id, name, date: date || "", status: "active", createdAt: now(), createdBy };
}
const closeIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>`;
const closeBtn = () =>
  `<button type="button" class="btn ghost sm icon" onclick="closeSheet()" aria-label="${t("close")}" title="${t("close")}">${closeIcon}</button>`;
const proofIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M21 16l-5-4-4 3-3-2-6 5"/></svg>`;
const downloadIcon = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>`;
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
    xls: "Unduh Excel",
    add: "+ Transaksi",
    net: "Saldo bersih",
    ticket: "Penjualan kursi",
    don: "Donasi & investor",
    exp: "Pengeluaran",
    wait: "Menunggu cek mutasi",
    trend: "Pemasukan {n} hari terakhir",
    seats: "Ketersediaan kursi",
    left: "kursi tersisa",
    soldOf: "{a} terjual dari {b}",
    queue: "Antre cek mutasi",
    noQueue: "Semua transaksi sudah diverifikasi.",
    topBuy: "Pembeli kursi terbanyak",
    topDon: "Donatur & investor terbesar",
    inRek: "Masuk",
    edit: "Ubah",
    all: "Semua",
    unchecked: "Belum dicek",
    seatsW: "Kursi",
    donW: "Donasi",
    expW: "Pengeluaran",
    date: "Tanggal",
    name: "Nama",
    detail: "Rincian",
    amount: "Jumlah",
    noTx: "Belum ada transaksi.",
    noTxSub: "Tekan + Transaksi untuk mencatat yang pertama.",
    newTx: "Catat transaksi",
    editTx: "Ubah transaksi",
    close: "Tutup",
    tBuy: "Pembelian kursi",
    tDon: "Donasi / investor",
    tExp: "Pengeluaran",
    bankTo: "Rekening penerima",
    bankFrom: "Sumber dana",
    nameBuy: "Nama pembeli",
    nameDon: "Nama donatur / investor",
    namePay: "Dibayarkan kepada",
    wa: "Nomor WhatsApp",
    catSeat: "Kategori kursi",
    catFree: "Kategori kursi gratis",
    nSeat: "Jumlah kursi",
    purpose: "Keperluan",
    amtIn: "Nominal transfer",
    amtOut: "Nominal pengeluaran",
    note: "Catatan bukti transfer",
    proof: "Bukti transfer (gambar)",
    proofBig: "Ukuran gambar maksimal 8 MB",
    proofHint: "Foto otomatis dikecilkan agar hemat penyimpanan",
    chooseFile: "Pilih berkas",
    noFileChosen: "Tidak ada berkas dipilih",
    fileAttached: "Gambar terpasang",
    viewProof: "Lihat bukti",
    searchTx: "Cari nama, telepon, catatan, rekening...",
    status: "Status uang",
    stW: "Menunggu cek mutasi",
    stV: "Sudah masuk rekening",
    save: "Simpan transaksi",
    saveGeneric: "Simpan",
    del: "Hapus",
    saved: "Transaksi tersimpan",
    deleted: "Transaksi dihapus",
    verified: "Ditandai sudah masuk rekening",
    needAmt: "Nominal belum diisi",
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
    addMethod: "+ Tambah metode",
    cashBreakdown: "Rincian pecahan uang",
    cashTotal: "Total tunai",
    chequeNo: "No. cek",
    chequeBank: "Bank penerbit",
    chequeDate: "Tanggal cair",
    saveSet: "Simpan pengaturan",
    setSaved: "Pengaturan disimpan",
    catTitle: "Kategori kursi",
    price: "Harga",
    quota: "Kuota",
    addCat: "+ Tambah kategori",
    dataT: "Data",
    dataP: "Data tersimpan bersama untuk semua pengguna aplikasi ini.",
    clearAll: "Hapus semua transaksi",
    cleared: "Semua transaksi dihapus",
    imp: "Impor data",
    impDrop: "Pilih berkas Excel atau CSV",
    impSub: "Gunakan template agar kolom terbaca otomatis",
    dlTpl: "Unduh template",
    tplTitle: "Kolom template",
    tplP: "Ubah nama kolom bila berkas Anda memakai istilah lain.",
    field: "Data",
    colName: "Nama kolom di berkas",
    prev: "Pratinjau impor",
    rowsOk: "{n} baris siap diimpor",
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
  },
  en: {
    dash: "Dashboard",
    tx: "Transactions",
    board: "Rankings",
    admin: "Admin",
    set: "Settings",
    users: "Users",
    logs: "Activity log",
    xls: "Export Excel",
    add: "+ Transaction",
    net: "Net balance",
    ticket: "Seat sales",
    don: "Donations & investors",
    exp: "Expenses",
    wait: "Awaiting bank check",
    trend: "Income, last {n} days",
    seats: "Seat availability",
    left: "seats left",
    soldOf: "{a} sold of {b}",
    queue: "Awaiting bank check",
    noQueue: "Everything is verified.",
    topBuy: "Top seat buyers",
    topDon: "Top donors & investors",
    inRek: "Received",
    edit: "Edit",
    all: "All",
    unchecked: "Unchecked",
    seatsW: "Seats",
    donW: "Donation",
    expW: "Expense",
    date: "Date",
    name: "Name",
    detail: "Details",
    amount: "Amount",
    noTx: "No transactions yet.",
    noTxSub: "Tap + Transaction to record the first one.",
    newTx: "Record transaction",
    editTx: "Edit transaction",
    close: "Close",
    tBuy: "Seat purchase",
    tDon: "Donation / investor",
    tExp: "Expense",
    bankTo: "Receiving account",
    bankFrom: "Paid from",
    nameBuy: "Buyer name",
    nameDon: "Donor / investor name",
    namePay: "Paid to",
    wa: "WhatsApp number",
    catSeat: "Seat category",
    catFree: "Free seat category",
    nSeat: "Number of seats",
    purpose: "Purpose",
    amtIn: "Transfer amount",
    amtOut: "Expense amount",
    note: "Transfer proof note",
    proof: "Transfer proof (image)",
    proofBig: "Image must be under 8 MB",
    proofHint: "Photos are automatically resized to save space",
    chooseFile: "Choose file",
    noFileChosen: "No file chosen",
    fileAttached: "Image attached",
    viewProof: "View proof",
    searchTx: "Search name, phone, note, account...",
    status: "Payment status",
    stW: "Awaiting bank check",
    stV: "Received in account",
    save: "Save transaction",
    saveGeneric: "Save",
    del: "Delete",
    saved: "Transaction saved",
    deleted: "Transaction deleted",
    verified: "Marked as received",
    needAmt: "Amount is required",
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
    addMethod: "+ Add method",
    cashBreakdown: "Cash denomination breakdown",
    cashTotal: "Cash total",
    chequeNo: "Cheque no.",
    chequeBank: "Issuing bank",
    chequeDate: "Clearing date",
    saveSet: "Save settings",
    setSaved: "Settings saved",
    catTitle: "Seat categories",
    price: "Price",
    quota: "Quota",
    addCat: "+ Add category",
    dataT: "Data",
    dataP: "Data is shared across everyone using this app.",
    clearAll: "Delete all transactions",
    cleared: "All transactions deleted",
    imp: "Import data",
    impDrop: "Choose an Excel or CSV file",
    impSub: "Use the template so columns are detected automatically",
    dlTpl: "Download template",
    tplTitle: "Template columns",
    tplP: "Rename a column if your file uses different wording.",
    field: "Data",
    colName: "Column name in file",
    prev: "Import preview",
    rowsOk: "{n} rows ready to import",
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
  bank: { id: "Rekening", en: "Account" },
  status: { id: "Status", en: "Status" },
  note: { id: "Catatan", en: "Note" },
};

/* ================= hitung ================= */
function sums(tx = S.tx, cats = D().cats) {
  const v = tx.filter((x) => x.status === "verified");
  const inTicket = v
    .filter((x) => x.type === "ticket")
    .reduce((a, b) => a + b.amount, 0);
  const inDon = v
    .filter((x) => x.type === "donation")
    .reduce((a, b) => a + b.amount, 0);
  const exp = v
    .filter((x) => x.type === "expense")
    .reduce((a, b) => a + b.amount, 0);
  const pd = tx.filter((x) => x.status === "pending" && x.type !== "expense");
  const quota = cats.reduce((a, c) => a + (+c.q || 0), 0);
  const sold = v
    .filter((x) => x.type !== "expense")
    .reduce((a, b) => a + (+b.seats || 0), 0);
  const held = pd.reduce((a, b) => a + (+b.seats || 0), 0);
  return {
    inTicket,
    inDon,
    exp,
    quota,
    sold,
    held,
    pendAmt: pd.reduce((a, b) => a + b.amount, 0),
    pendN: pd.length,
    net: inTicket + inDon - exp,
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
function byPerson(ty, tx = S.tx) {
  const m = {};
  tx
    .filter(
      (x) =>
        x.status === "verified" && (ty ? x.type === ty : x.type !== "expense"),
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
  if (S.config.methods) return;
  const methods = methodsFromBanks(S.config.banks);
  await mutateEvent(() => {
    S.config.methods = methods;
    delete S.config.banks;
  }, null);
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
    if (tab === "dash") drawChart();
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
    <div class="mark">TS</div><h1 style="font-size:19px">${t("setupT")}</h1></div>
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
    <div class="mark">TS</div><div style="min-width:0"><h1 style="font-size:19px">${f === "forgot" ? t("forgotT") : t("welcome")}</h1>
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
    <div class="mark">TS</div><div style="min-width:0"><h1 style="font-size:19px">${t("chooseEvent")}</h1>
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
    <div class="mark">TS</div>
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
function vHubEvents() {
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
  const si = (k) =>
    hubEventsSort.k === k ? (hubEventsSort.dir === "asc" ? " ▲" : " ▼") : "";
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
      <input style="max-width:220px" placeholder="${t("searchTx")}" value="${esc(hubEventsQ)}" oninput="onHubEventsSearch(this.value)">
      <button class="btn sm" onclick="openNewEvent()">${t("newEvent")}</button></div></div>
    <div style="overflow-x:auto"><table><thead><tr>
      <th class="sortable" onclick="setHubEventsSort('name')">${t("name")}${si("name")}</th>
      <th class="sortable" onclick="setHubEventsSort('date')">${t("evDate")}${si("date")}</th>
      <th class="sortable" onclick="setHubEventsSort('status')">${t("status")}${si("status")}</th>
      <th></th></tr></thead><tbody>
    ${
      items.length
        ? items
            .map(
              (e) => `<tr>
      <td style="font-weight:600">${esc(e.name)}</td>
      <td class="hint mono">${e.date || "-"}</td>
      <td><span class="tag ${e.status === "active" ? "t-ok" : "t-exp"}">${e.status === "active" ? t("active") : t("archived")}</span></td>
      <td style="text-align:right;white-space:nowrap">
        ${e.status === "active" ? `<button class="btn ghost sm" onclick="switchEvent('${e.id}')">${t("enterWorkspace")}</button> ` : ""}
        <button class="btn ghost sm" onclick="toggleArchive('${e.id}')">${e.status === "active" ? t("archive") : t("restore")}</button>
      </td></tr>`,
            )
            .join("")
        : `<tr><td colspan="4" class="empty">${t("noneYet")}</td></tr>`
    }
    </tbody></table></div>${pagerHtml(page, totalPages, "setHubEventsPage")}</div>`;
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
    <div style="overflow-x:auto"><table><thead><tr><th>${t("name")}</th><th>${t("events")}</th><th>${t("recorded")}</th><th>${t("verified")}</th></tr></thead><tbody>
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
  const tabs = [
    ["dash", t("dash")],
    ["tx", t("tx")],
    ["board", t("board")],
  ];
  if (isAdmin()) tabs.push(["admin", t("admin")]);
  if (tab === "admin" && !isAdmin()) tab = "dash";
  return `${
    imp
      ? `<div class="banner">${t("impBanner", { n: esc(imp.name) })}
      <button class="btn sm" onclick="stopImp()">${t("backAdmin")}</button></div>`
      : ""
  }
  <header><div class="wrap">
    <div class="mark">TS</div>
    <div style="flex:1;min-width:0"><h1 style="font-size:17px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(D().event || "Konser")}</h1>
      <div class="sub">${(D().date ? new Date(D().date + "T00:00:00").toLocaleDateString(lang === "id" ? "id-ID" : "en-GB", { day: "numeric", month: "short", year: "numeric" }) + " · " : "") + s.sold + " " + t("seatsSold")}</div></div>
    ${themeBtn()}
    ${langSeg()}
    <button class="btn ghost sm xls-btn" onclick="openExportPeriod()">${downloadIcon}<span class="xls-label">${t("xls")}</span></button>
    <div class="avatar" onclick="openMe()" title="${esc(a.name)}">${esc(a.name.slice(0, 1).toUpperCase())}</div>
  </div></header>
  <div class="wrap">${{ dash: vDash, tx: vTx, board: vBoard, admin: vAdmin }[tab]()}</div>
  ${tab === "tx" || tab === "dash" ? `<button class="btn fab" onclick="openTx()">${t("add")}</button>` : ""}
  <nav>${tabs.map(([k, l]) => `<button class="${tab === k ? "on" : ""}" onclick="go('${k}')"><span class="dot"></span>${l}</button>`).join("")}</nav>`;
}
function go(k) {
  tab = k;
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
  window._lq = setTimeout(render, 250);
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
  window._heq = setTimeout(render, 250);
}
function setHubStaffPage(p) {
  hubStaffPage = p;
  render();
}
function onHubStaffSearch(v) {
  hubStaffQ = v;
  hubStaffPage = 1;
  clearTimeout(window._hsq);
  window._hsq = setTimeout(render, 250);
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
    <span class="tag ${a.role === "admin" ? "t-adm" : "t-tic"}">${a.role === "admin" ? t("admins") : t("treas")}</span>
    ${
      imp
        ? `<p class="hint" style="margin-top:12px">${t("byAdmin", { n: esc(me.name) })}</p>
      <button class="btn wide" style="margin-top:8px" onclick="stopImp()">${t("backAdmin")}</button>`
        : ""
    }
    ${
      events.length
        ? `<div class="field" style="margin-top:14px"><label>${t("switchEvent")}</label>
      ${events
        .map(
          (e) =>
            `<button class="btn ghost wide" style="display:flex;justify-content:flex-start;text-align:left;margin-bottom:6px" onclick="switchEvent('${e.id}')">${esc(e.name)}</button>`,
        )
        .join("")}</div>`
        : ""
    }
    ${
      isAdmin() && screen === "app"
        ? `<button class="btn ghost wide" style="margin-top:6px" onclick="goHub()">${t("adminConsole")}</button>`
        : ""
    }
    <button class="btn danger wide" style="margin-top:14px" onclick="closeSheet();doSignOut()">${t("signOut")}</button>`);
}
async function stopImp() {
  const n = imp.name;
  imp = null;
  closeSheet();
  await saveSession();
  render();
  toast(t("backAdmin") + " · " + n);
}

/* ================= dashboard ================= */
function pSeats() {
  const s = sums(),
    pct = s.quota ? Math.round((s.sold / s.quota) * 100) : 0;
  return `<span class="label">${t("seats")}</span>
  <div class="mono" style="font-size:29px;font-weight:700;letter-spacing:-.035em">${s.avail}<span style="font-size:15px;color:var(--tx2);font-weight:500"> ${t("left")}</span></div>
  <div class="bar"><i style="width:${Math.min(pct, 100)}%"></i></div>
  <div class="rowsp hint" style="flex:none"><span>${t("soldOf", { a: s.sold, b: s.quota })}</span><span>${pct}%</span></div>
  <div class="scroll" style="margin-top:6px">${D()
    .cats.map((c) => {
      const sd = S.tx
        .filter((x) => x.status === "verified" && x.cat === c.n)
        .reduce((a, b) => a + (+b.seats || 0), 0);
      return `<div class="rowsp" style="padding:7px 0;border-top:1px solid var(--line2);font-size:15px">
      <span>${esc(c.n)} <span class="hint">${rp(c.p)}</span></span><span class="mono" style="font-weight:600">${sd}/${c.q}</span></div>`;
    })
    .join("")}</div>`;
}
function pQueue() {
  const s = sums(),
    q = S.tx.filter((x) => x.status === "pending");
  return `<div class="rowsp" style="flex:none"><span class="label">${t("queue")}</span>
    <span class="mono amb" style="font-weight:700;font-size:17px">${rpk(s.pendAmt)}</span></div>
  <div class="scroll" style="flex:1;margin-top:4px">${
    q
      .map(
        (
          x,
        ) => `<div class="rowsp" style="padding:7px 0;border-top:1px solid var(--line2)">
    <div style="min-width:0"><div style="font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(x.name)}</div>
    <div class="hint mono">${rpk(x.amount)} · ${esc(x.bank || "-")}</div></div>
    <button class="btn ok sm" onclick="verify('${x.id}')">${t("inRek")}</button></div>`,
      )
      .join("") ||
    `<div class="hint" style="padding-top:10px">${t("noQueue")}</div>`
  }</div>`;
}
function pRank() {
  const a = byPerson("ticket").slice(0, 5),
    b = byPerson("donation").slice(0, 5);
  const row = (x, i) =>
    `<div class="rowsp" style="padding:5px 0;font-size:15px"><span style="min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${i + 1}. ${esc(x.name)}</span><span class="mono" style="font-weight:600">${rpk(x.amount)}</span></div>`;
  return `<span class="label">${t("topBuy")}</span><div style="margin-bottom:8px">${a.map(row).join("") || `<div class="hint">${t("noneYet")}</div>`}</div>
  <span class="label">${t("topDon")}</span><div class="scroll">${b.map(row).join("") || `<div class="hint">${t("noneYet")}</div>`}</div>`;
}
function vDash() {
  const s = sums(),
    nw = narrow();
  const k = (l, v, c) =>
    `<div class="card kpi act" style="padding:11px 13px;justify-content:center"><div class="label">${l}</div><div class="n mono ${c || ""}">${v}</div></div>`;
  const kpis =
    k(t("net"), rpk(s.net), s.net >= 0 ? "pos" : "neg") +
    k(t("ticket"), rpk(s.inTicket), "blu") +
    k(t("don"), rpk(s.inDon)) +
    k(t("exp"), rpk(s.exp), "neg") +
    (nw ? "" : k(t("wait"), rpk(s.pendAmt), "amb"));
  const chart = `<div class="card" style="grid-row:span ${nw ? 1 : 2}">
      <div class="rowsp" style="flex:none"><span class="label">${t("trend", { n: nw ? 7 : 14 })}</span>
      <span class="mono" style="font-weight:700;font-size:17px">${rpk(s.inTicket + s.inDon)}</span></div>
      <div class="chartbox"><canvas id="c1"></canvas></div></div>`;
  if (nw)
    return `<div class="fit">
    <div class="grid" style="grid-template-columns:1fr 1fr;flex:none">${kpis}</div>${chart}
    <div class="card" style="flex:1.1;gap:6px">
      <div class="seg" style="align-self:flex-start;flex:none">${[
        ["seats", t("seatsW")],
        ["queue", t("queue")],
        ["rank", t("rank")],
      ]
        .map(
          ([p, l]) =>
            `<button class="${panel === p ? "on" : ""}" onclick="setPanel('${p}')">${l}</button>`,
        )
        .join("")}</div>
      <div style="display:flex;flex-direction:column;flex:1;min-height:0">${{ seats: pSeats, queue: pQueue, rank: pRank }[panel]()}</div>
    </div></div>`;
  return `<div class="fit">
    <div class="grid" style="grid-template-columns:repeat(5,1fr);flex:none">${kpis}</div>
    <div class="grid" style="grid-template-columns:1.45fr 1fr 1fr;grid-template-rows:1fr 1fr;flex:1;min-height:0">
      ${chart}<div class="card">${pSeats()}</div><div class="card">${pQueue()}</div>
      <div class="card" style="grid-column:span 2">${pRank()}</div></div></div>`;
}
function drawChart() {
  const el = document.getElementById("c1");
  if (!el) return;
  const s = seriesByDay(narrow() ? 7 : 14);
  const cs = getComputedStyle(document.documentElement);
  const cv = (n) => cs.getPropertyValue(n).trim();
  chart1 && chart1.destroy();
  chart1 = new Chart(el, {
    type: "bar",
    data: {
      labels: s.map((x) => x.label),
      datasets: [
        {
          data: s.map((x) => x.v),
          backgroundColor: cv("--blue"),
          hoverBackgroundColor: cv("--blue-d"),
          borderRadius: 5,
          maxBarThickness: 38,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 500 },
      layout: { padding: { top: 6 } },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (c) => rp(c.parsed.y) } },
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
          ticks: {
            color: cv("--tx2"),
            font: { size: 12 },
            maxTicksLimit: 4,
            callback: (v) =>
              v >= 1e6
                ? v / 1e6 + (lang === "id" ? " jt" : " M")
                : v >= 1000
                  ? v / 1000 + (lang === "id" ? " rb" : " k")
                  : v,
          },
        },
      },
    },
  });
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
function onTxSearch(v) {
  txQ = v;
  txPage = 1;
  clearTimeout(window._tq);
  window._tq = setTimeout(render, 250);
}
function vTx() {
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
  const si = (k) =>
    txSort.k === k ? (txSort.dir === "asc" ? " ▲" : " ▼") : "";
  return `<div style="padding:12px 0 110px">
  <div class="rowsp" style="margin-bottom:11px;flex-wrap:wrap;gap:8px">
    <div class="chips">${[
      ["all", t("all")],
      ["pending", t("unchecked")],
      ["ticket", t("seatsW")],
      ["donation", t("donW")],
      ["expense", t("expW")],
    ]
      .map(
        ([k, l]) =>
          `<button class="chip ${filter === k ? "on" : ""}" onclick="setFilter('${k}')">${l}</button>`,
      )
      .join("")}</div>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;min-width:200px">
      <input style="max-width:280px" placeholder="${t("searchTx")}" value="${esc(txQ)}" oninput="onTxSearch(this.value)">
      <button class="btn ghost sm" onclick="openImport()">${t("imp")}</button></div></div>
  <div class="card" style="padding:2px 10px">${
    f.length
      ? `<div style="overflow-x:auto"><table><thead><tr>
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
          ? `<span class="tag t-exp">${t("expW")}</span> ${esc(x.cat || "")}`
          : `<span class="tag ${x.type === "donation" ? "t-ok" : "t-tic"}">${x.type === "donation" ? t("donW") : t("seatsW")}</span> ${esc(x.cat || "")} · ${x.seats || 0}`
      }
        <div class="hint">${esc(x.bank || "-")}${x.chequeNo ? " · " + t("chequeNo") + " " + esc(x.chequeNo) : ""}${x.note ? " · " + esc(x.note) : ""}</div></td>
      <td class="mono ${x.type === "expense" ? "neg" : ""}" style="text-align:right;white-space:nowrap;font-weight:600">${x.type === "expense" ? "−" : ""}${rp(x.amount)}
        <div style="margin-top:3px">${x.status === "verified" ? `<span class="tag t-ok">${t("inRek")}</span>` : `<span class="tag t-wait">${t("stW")}</span>`}</div></td>
      <td style="text-align:right;white-space:nowrap">${x.proof ? `<button class="btn ghost sm icon" onclick="viewProofById('${x.id}')" title="${t("viewProof")}">${proofIcon}</button> ` : ""}${x.status === "pending" ? `<button class="btn ok sm" onclick="verify('${x.id}')">${t("inRek")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openTx('${x.id}')">${t("edit")}</button></td></tr>`,
      )
      .join("")}
    </tbody></table></div>${pagerHtml(page, totalPages, "setTxPage")}`
      : `<div class="empty"><b>${t("noTx")}</b><br>${t("noTxSub")}</div>`
  }</div></div>`;
}
async function verify(id) {
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
  const x = id
    ? S.tx.find((v) => v.id === id)
    : {
        id: "",
        type: "ticket",
        date: today(),
        name: "",
        phone: "",
        cat: D().cats[0]?.n || "",
        seats: 1,
        amount: D().cats[0]?.p || 0,
        bank: D().methods[0]?.name || "",
        note: "",
        proof: "",
        status: "pending",
      };
  sheet(`<div class="rowsp" style="margin-bottom:14px"><h2 style="font-size:20px">${id ? t("editTx") : t("newTx")}</h2>
      ${closeBtn()}</div>
    <div class="chips" style="margin-bottom:14px" id="typeChips">
      ${[
        ["ticket", t("tBuy")],
        ["donation", t("tDon")],
        ["expense", t("tExp")],
      ]
        .map(
          ([k, l]) =>
            `<button class="chip ${x.type === k ? "on" : ""}" data-k="${k}" onclick="setType('${k}')">${l}</button>`,
        )
        .join("")}</div>
    <input type="hidden" id="f_id" value="${x.id}"><input type="hidden" id="f_type" value="${x.type}">
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field"><label>${t("date")}</label><input type="date" id="f_date" value="${x.date}"></div>
      <div class="field"><label id="l_bank">${t("bankTo")}</label><select id="f_bank" onchange="onMethodChange()">${D()
        .methods.map(
          (m) =>
            `<option value="${esc(m.name)}" data-type="${m.type}" ${m.name === x.bank ? "selected" : ""}>${esc(m.name)}</option>`,
        )
        .join("")}</select></div>
      <div class="field" id="w_name"><label id="l_name">${t("nameBuy")}</label><input id="f_name" value="${esc(x.name)}"></div>
      <div class="field" id="w_phone"><label>${t("wa")}</label><input id="f_phone" inputmode="tel" value="${esc(x.phone)}" placeholder="0812..."></div>
      <div class="field" id="w_cat"><label id="l_cat">${t("catSeat")}</label><select id="f_cat" onchange="recalc()">${D()
        .cats.map(
          (c) =>
            `<option ${c.n === x.cat ? "selected" : ""}>${esc(c.n)}</option>`,
        )
        .join("")}</select></div>
      <div class="field" id="w_seats"><label>${t("nSeat")}</label>
        <div class="stepper">
          <button type="button" class="step-btn" onclick="stepSeats(-1)" aria-label="-">−</button>
          <input type="number" inputmode="numeric" min="0" id="f_seats" value="${x.seats || 0}" oninput="recalc()">
          <button type="button" class="step-btn" onclick="stepSeats(1)" aria-label="+">+</button>
        </div></div>
    </div>
    <div class="field" id="w_expcat" style="display:none">
      <label>${t("purpose")}</label><input id="f_expcat" value="${x.type === "expense" ? esc(x.cat) : ""}">
    </div>
    <div class="field" id="w_cash" style="display:none">
      <label>${t("cashBreakdown")}</label>
      <div class="grid" style="grid-template-columns:repeat(3,1fr)">${CASH_DENOMS.map(
        (d) =>
          `<div class="field"><label>${rp(d)}</label><input type="number" min="0" inputmode="numeric" class="cash-denom" data-denom="${d}" value="${(x.cashBreakdown && x.cashBreakdown[d]) || 0}" oninput="recalcCash()"></div>`,
      ).join("")}</div>
    </div>
    <div class="grid" id="w_cheque" style="display:none;grid-template-columns:1fr 1fr 1fr">
      <div class="field"><label>${t("chequeNo")}</label><input id="f_chequeNo" value="${esc(x.chequeNo || "")}"></div>
      <div class="field"><label>${t("chequeBank")}</label><input id="f_chequeBank" value="${esc(x.chequeBank || "")}"></div>
      <div class="field"><label>${t("chequeDate")}</label><input type="date" id="f_chequeDate" value="${x.chequeDate || ""}"></div>
    </div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field amt-field"><label id="l_amount">${t("amtIn")}</label><input type="number" inputmode="numeric" id="f_amount" value="${x.amount}" oninput="prev()">
        <div class="hint mono" id="prev" style="margin-top:4px">${rp(x.amount)}</div></div>
      <div class="field"><label>${t("status")}</label><select id="f_status">
        <option value="pending" ${x.status === "pending" ? "selected" : ""}>${t("stW")}</option>
        <option value="verified" ${x.status === "verified" ? "selected" : ""}>${t("stV")}</option></select></div></div>
    <div class="field"><label>${t("note")}</label><textarea id="f_note" rows="2">${esc(x.note)}</textarea></div>
    <div class="field"><label>${t("proof")}</label>
      <div class="filepick">
        <button type="button" class="btn ghost sm" onclick="document.getElementById('f_proofFile').click()">${t("chooseFile")}</button>
        <span class="hint" id="proofFileName">${x.proof ? t("fileAttached") : t("noFileChosen")}</span>
      </div>
      <input type="file" id="f_proofFile" accept="image/*" style="display:none" onchange="onProofChange(event)">
      <div class="hint" style="margin-top:4px">${t("proofHint")}</div>
      <div class="proof-prev" id="proofPrev" style="margin-top:8px">${x.proof ? `<img src="${x.proof}" onclick="viewProofPreview()">` : ""}</div>
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
    <div class="rowsp" style="margin-top:14px">${id ? `<button class="btn danger" onclick="delTx('${id}')">${t("del")}</button>` : "<span></span>"}
      <button class="btn" onclick="saveTx()">${t("save")}</button></div>`);
  setType(x.type, true);
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
function setType(k, init) {
  document.getElementById("f_type").value = k;
  document
    .querySelectorAll("#typeChips .chip")
    .forEach((c) => c.classList.toggle("on", c.dataset.k === k));
  const e = k === "expense";
  ["w_cat", "w_seats"].forEach(
    (i) => (document.getElementById(i).style.display = e ? "none" : "block"),
  );
  document.getElementById("w_phone").style.display = e ? "none" : "block";
  // pengeluaran menyembunyikan phone/cat/seats, jadi field nama jadi sendirian
  // di barisnya - bentangkan penuh 2 kolom biar tidak kelihatan terpotong separuh
  document.getElementById("w_name").style.gridColumn = e ? "1 / -1" : "";
  document.getElementById("w_expcat").style.display = e ? "block" : "none";
  document.getElementById("l_name").textContent = e
    ? t("namePay")
    : k === "donation"
      ? t("nameDon")
      : t("nameBuy");
  document.getElementById("l_bank").textContent = e
    ? t("bankFrom")
    : t("bankTo");
  document.getElementById("l_amount").textContent = e
    ? t("amtOut")
    : t("amtIn");
  document.getElementById("l_cat").textContent =
    k === "donation" ? t("catFree") : t("catSeat");
  if (!init && k === "ticket") recalc();
  onMethodChange();
}
function stepSeats(d) {
  const el = document.getElementById("f_seats");
  el.value = Math.max(0, (+el.value || 0) + d);
  recalc();
}
function recalc() {
  if (document.getElementById("f_type").value !== "ticket") return;
  const c = D().cats.find(
    (x) => x.n === document.getElementById("f_cat").value,
  );
  if (c) {
    document.getElementById("f_amount").value =
      (+document.getElementById("f_seats").value || 0) * c.p;
    prev();
  }
}
function currentMethodType() {
  const sel = document.getElementById("f_bank");
  return sel?.selectedOptions[0]?.dataset.type || "bank";
}
function onMethodChange() {
  const type = currentMethodType();
  document.getElementById("w_cash").style.display = type === "cash" ? "block" : "none";
  document.getElementById("w_cheque").style.display = type === "cheque" ? "grid" : "none";
  const amt = document.getElementById("f_amount");
  amt.readOnly = type === "cash";
  if (type === "cash") recalcCash();
}
function recalcCash() {
  let total = 0;
  document
    .querySelectorAll(".cash-denom")
    .forEach((inp) => (total += (+inp.dataset.denom || 0) * (+inp.value || 0)));
  document.getElementById("f_amount").value = total;
  prev();
}
async function saveTx() {
  const g = (i) => document.getElementById("f_" + i).value,
    ty = g("type"),
    id = g("id"),
    methodType = currentMethodType();
  const x = {
    id: id || uid(),
    type: ty,
    date: g("date") || today(),
    name: g("name").trim() || "—",
    phone: ty === "expense" ? "" : g("phone"),
    cat:
      ty === "expense" ? document.getElementById("f_expcat").value : g("cat"),
    seats: ty === "expense" ? 0 : +g("seats") || 0,
    amount: +g("amount") || 0,
    bank: g("bank"),
    methodType,
    cashBreakdown:
      methodType === "cash"
        ? Object.fromEntries(
            Array.from(document.querySelectorAll(".cash-denom"))
              .filter((i) => +i.value > 0)
              .map((i) => [i.dataset.denom, +i.value]),
          )
        : {},
    chequeNo: methodType === "cheque" ? g("chequeNo") : "",
    chequeBank: methodType === "cheque" ? g("chequeBank") : "",
    chequeDate: methodType === "cheque" ? g("chequeDate") : "",
    note: g("note"),
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
  if (!x.amount) return toast(t("needAmt"));
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
  const {
    items: a,
    page: pageBuy,
    totalPages: totalBuy,
  } = paginate(byPerson("ticket"), boardBuyPage, BOARD_SIZE);
  const {
    items: b,
    page: pageDon,
    totalPages: totalDon,
  } = paginate(byPerson("donation"), boardDonPage, BOARD_SIZE);
  const row = (
    x,
    rank,
    u,
  ) => `<div class="lb" style="animation-delay:${((rank - 1) % BOARD_SIZE) * 35}ms"><div class="rank ${rank === 1 ? "g1" : ""}">${rank}</div>
    <div style="flex:1;min-width:0"><div style="font-weight:600">${esc(x.name)}</div>
    <div class="hint">${x.n} ${t("trans")} · ${x.seats} ${t("seatsW").toLowerCase()}${u}</div></div>
    <div class="mono" style="font-weight:700">${rp(x.amount)}</div></div>`;
  return `<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(330px,1fr));padding:12px 0 110px">
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("topBuy")}</h2>
      ${a.map((x, i) => row(x, (pageBuy - 1) * BOARD_SIZE + i + 1, "")).join("") || `<div class="empty">${t("noneYet")}</div>`}
      ${pagerHtml(pageBuy, totalBuy, "setBoardBuyPage")}</div>
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("topDon")}</h2>
      ${b.map((x, i) => row(x, (pageDon - 1) * BOARD_SIZE + i + 1, " " + t("free"))).join("") || `<div class="empty">${t("noneYet")}</div>`}
      ${pagerHtml(pageDon, totalDon, "setBoardDonPage")}</div></div>`;
}

/* ================= admin ================= */
function vAdmin() {
  if (!isAdmin()) return `<div class="empty">${t("onlyAdmin")}</div>`;
  return `<div style="padding:12px 0 110px">
    <div class="seg" style="margin-bottom:12px">${[
      ["logs", t("logs")],
      ["set", t("set")],
    ]
      .map(
        ([k, l]) =>
          `<button class="${adminTab === k ? "on" : ""}" onclick="setAdminTab('${k}')">${l}</button>`,
      )
      .join("")}</div>
    ${{ logs: vLogs, set: vSet }[adminTab]()}</div>`;
}
function vHubStaff() {
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
  const si = (k) =>
    hubStaffSort.k === k ? (hubStaffSort.dir === "asc" ? " ▲" : " ▼") : "";
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("staff")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;flex-wrap:wrap;min-width:220px">
      <div class="chips">${[
        ["all", t("all")],
        ["admin", t("admins")],
        ["treasurer", t("treas")],
      ]
        .map(
          ([k2, l]) =>
            `<button class="chip ${hubStaffFilter === k2 ? "on" : ""}" onclick="setHubStaffFilter('${k2}')">${l}</button>`,
        )
        .join("")}</div>
      <input style="max-width:220px" placeholder="${t("searchTx")}" value="${esc(hubStaffQ)}" oninput="onHubStaffSearch(this.value)">
      <button class="btn sm" onclick="openUser()">${t("addUser")}</button></div></div>
  <div style="overflow-x:auto"><table><thead><tr>
    <th class="sortable" onclick="setHubStaffSort('name')">${t("name")}${si("name")}</th>
    <th class="sortable" onclick="setHubStaffSort('role')">${t("role")}${si("role")}</th>
    <th class="sortable" onclick="setHubStaffSort('active')">${t("status")}${si("active")}</th>
    <th class="sortable" onclick="setHubStaffSort('last')">${t("lastIn")}${si("last")}</th><th></th></tr></thead><tbody>
    ${
      items.length
        ? items
            .map(
              (u) => `<tr>
      <td><div style="font-weight:600">${esc(u.name)}</div><div class="hint">${esc(u.email)}${u.provider === "google" ? " · Google" : ""}</div></td>
      <td><span class="tag ${u.role === "admin" ? "t-adm" : "t-tic"}">${u.role === "admin" ? t("admins") : t("treas")}</span></td>
      <td><span class="tag ${u.active ? "t-ok" : "t-exp"}">${u.active ? t("active") : t("inactive")}</span></td>
      <td class="hint mono">${u.last ? new Date(u.last).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : t("never")}</td>
      <td style="text-align:right;white-space:nowrap">
        ${u.email !== me.email && u.active ? `<button class="btn ghost sm" onclick="startImp('${esc(u.email)}')">${t("loginAs")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openUser('${esc(u.email)}')">${t("edit")}</button></td></tr>`,
            )
            .join("")
        : `<tr><td colspan="5" class="empty">${t("noneYet")}</td></tr>`
    }
  </tbody></table></div>${pagerHtml(page, totalPages, "setHubStaffPage")}</div>`;
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
    `${old ? t("userSaved") : t("userAdded")}: ${n} (${role === "admin" ? t("admins") : t("treas")})`,
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
function vLogs() {
  const q = logQ.toLowerCase();
  const all = S.logs.filter(
    (l) =>
      !q ||
      [l.user, l.email, t(l.act), l.det].join(" ").toLowerCase().includes(q),
  );
  const { items: rows, page, totalPages } = paginate(all, logPage);
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("logs")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;min-width:220px">
      <input style="max-width:320px" placeholder="${t("searchLog")}" value="${esc(logQ)}"
        oninput="onLogSearch(this.value)">
      <button class="btn ghost sm" onclick="exportLogs()">${t("exportLog")}</button></div></div>
  ${
    rows.length
      ? `<div style="overflow-x:auto"><table><thead><tr>
    <th>${t("time")}</th><th>${t("user")}</th><th>${t("action")}</th><th>${t("info")}</th></tr></thead><tbody>
    ${rows
      .map(
        (l) => `<tr>
      <td class="hint mono" style="white-space:nowrap">${new Date(l.ts).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</td>
      <td><div style="font-weight:600">${esc(l.user)}</div><div class="hint">${esc(l.email)}${l.asAdmin ? " · " + t("byAdmin", { n: esc(l.asAdmin) }) : ""}</div></td>
      <td><span class="tag ${/Delete|Hapus/.test(t(l.act)) ? "t-exp" : /Verif/.test(t(l.act)) ? "t-ok" : "t-tic"}">${t(l.act)}</span></td>
      <td class="hint">${esc(l.det)}</td></tr>`,
      )
      .join("")}</tbody></table></div>${pagerHtml(page, totalPages, "setLogPage")}`
      : `<div class="empty">${t("noLog")}</div>`
  }</div>`;
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
      <div style="overflow-x:auto"><table><thead><tr><th>${t("name")}</th><th>${t("price")}</th><th>${t("quota")}</th><th></th></tr></thead><tbody>
      ${D()
        .cats.map(
          (
            c,
            i,
          ) => `<tr><td><input value="${esc(c.n)}" onchange="editCat(${i},'n',this.value)"></td>
        <td><input type="number" value="${c.p}" onchange="editCat(${i},'p',+this.value)"></td>
        <td><input type="number" value="${c.q}" onchange="editCat(${i},'q',+this.value)"></td>
        <td><button class="btn danger sm" onclick="delCat(${i})">${t("del")}</button></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div><button class="btn ghost sm" style="margin-top:10px" onclick="addCat()">${t("addCat")}</button></div></div>
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("methods")}</h2>
      <div style="overflow-x:auto"><table><thead><tr><th>${t("methodName")}</th><th>${t("methodType")}</th><th></th></tr></thead><tbody>
      ${D()
        .methods.map(
          (m, i) => `<tr><td><input value="${esc(m.name)}" onchange="editMethod(${i},'name',this.value)"></td>
        <td><select onchange="editMethod(${i},'type',this.value)">
          <option value="bank" ${m.type === "bank" ? "selected" : ""}>${t("methodBank")}</option>
          <option value="cash" ${m.type === "cash" ? "selected" : ""}>${t("methodCash")}</option>
          <option value="cheque" ${m.type === "cheque" ? "selected" : ""}>${t("methodCheque")}</option>
          <option value="other" ${m.type === "other" ? "selected" : ""}>${t("methodOther")}</option>
        </select></td>
        <td><button class="btn danger sm" onclick="delMethod(${i})">${t("del")}</button></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div><button class="btn ghost sm" style="margin-top:10px" onclick="addMethod()">${t("addMethod")}</button></div></div>
    <div class="card"><h2 style="font-size:17px">${t("tplTitle")}</h2>
      <p class="hint" style="margin:4px 0 8px">${t("tplP")}</p>
      <div style="overflow-x:auto"><table><thead><tr><th>${t("field")}</th><th>${t("colName")}</th><th></th></tr></thead><tbody>
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
      S.config.cats.push({ n: "—", p: 0, q: 0 });
    },
    "actSet",
    t("addCat"),
  );
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
  if (c.f === "type")
    return { ticket: t("tBuy"), donation: t("tDon"), expense: t("tExp") }[x.type];
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
async function dlTemplate() {
  const id = lang === "id";
  const ex = [
    {
      type: "ticket",
      name: "Michael Jonathan",
      phone: "081234567890",
      cat: D().cats[0]?.n || "Reguler",
      seats: 2,
      amount: (D().cats[0]?.p || 0) * 2,
      bank: D().methods[0]?.name || "BCA",
      status: "verified",
      note: "Transfer 08.41",
      date: today(),
      custom: {},
    },
    {
      type: "donation",
      name: "PT Nada Jaya",
      phone: "",
      cat: D().cats[0]?.n || "Reguler",
      seats: 10,
      amount: 25000000,
      bank: D().methods[0]?.name || "BCA",
      status: "pending",
      note: "",
      date: today(),
      custom: {},
    },
    {
      type: "expense",
      name: "Sound System",
      phone: "",
      cat: "",
      seats: 0,
      amount: 7500000,
      bank: D().methods[0]?.name || "BCA",
      status: "verified",
      note: "DP 50%",
      date: today(),
      custom: {},
    },
  ];
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(id ? "Template" : "Template");
  addReportTitle(ws, periodLabel(today(), "week"));
  addTxTable(ws, ex, "TabelTemplate");
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
function normDate(v) {
  if (!v) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  if (/^\d+$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + +s * 864e5);
    return d.toISOString().slice(0, 10);
  }
  const d = new Date(s);
  return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}
function normType(v) {
  const s = String(v || "").toLowerCase();
  if (/don|invest/.test(s)) return "donation";
  if (/exp|keluar|biaya|beban/.test(s)) return "expense";
  return "ticket";
}
function normNum(v) {
  return +String(v ?? "").replace(/[^\d.-]/g, "") || 0;
}
// file yang di-download dari app sendiri (template maupun laporan) punya 3
// baris judul (LAPORAN KEUANGAN / periode / kosong) sebelum baris header -
// cari baris mana yang paling cocok dengan nama kolom di D().tpl, jangan
// asumsikan header selalu di baris pertama (juga tetap benar untuk file
// polos yang headernya memang di baris pertama)
function findHeaderRowIndex(rows) {
  const tplHeaders = D().tpl.map((c) => String(c.h).trim().toLowerCase());
  let bestIdx = 0,
    bestScore = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = (rows[i] || []).map((h) => String(h).trim().toLowerCase());
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
  const head = (rows[headerIdx] || []).map((h) => String(h).trim().toLowerCase()),
    idx = {};
  D().tpl.forEach((c) => {
    const i = head.indexOf(String(c.h).trim().toLowerCase());
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
    const debit = normNum(get(r, "debit")),
      credit = normNum(get(r, "credit")),
      legacyAmt = normNum(get(r, "amount")),
      amt = debit || credit || legacyAmt,
      dt = normDate(get(r, "date")) || today();
    if (!amt) {
      bad++;
      return;
    }
    let ty = normType(get(r, "type"));
    if (!String(get(r, "type") || "").trim() && credit && !debit) ty = "expense";
    const st = /masuk|verif|lunas|received|paid|ok/i.test(
      String(get(r, "status")),
    )
      ? "verified"
      : "pending";
    const custom = {};
    customCols.forEach((c) => {
      custom[c.f] = String(get(r, c.f) || "");
    });
    pendingImp.push({
      id: uid(),
      type: ty,
      date: dt,
      name: String(get(r, "name") || "—").trim(),
      phone: String(get(r, "phone") || ""),
      cat: String(
        get(r, "cat") || (ty === "expense" ? "" : D().cats[0]?.n || ""),
      ),
      seats: ty === "expense" ? 0 : normNum(get(r, "seats")),
      amount: amt,
      bank: String(get(r, "bank") || D().methods[0]?.name || ""),
      methodType:
        D().methods.find((m) => m.name === String(get(r, "bank") || ""))?.type ||
        "bank",
      cashBreakdown: {},
      chequeNo: "",
      chequeBank: "",
      chequeDate: "",
      note: String(get(r, "note") || ""),
      proof: "",
      status: st,
      by: acting().name,
      custom,
    });
  });
  const lbl = {
    ticket: t("seatsW"),
    donation: t("donW"),
    expense: t("expW"),
  };
  body.innerHTML = `<div style="margin-top:12px"><div class="rowsp"><span class="label">${t("prev")}</span>
    <span class="hint">${t("rowsOk", { n: pendingImp.length })}${bad ? " · " + t("rowsBad", { n: bad }) : ""}</span></div>
    <div class="scroll" style="max-height:230px;border:1px solid var(--line);border-radius:10px;margin-top:6px;padding:0 10px">
    <table><thead><tr><th>${t("date")}</th><th>${t("name")}</th><th>${t("detail")}</th><th style="text-align:right">${t("amount")}</th></tr></thead><tbody>
    ${pendingImp
      .map(
        (
          x,
        ) => `<tr><td class="mono hint">${x.date.slice(5)}</td><td style="font-weight:600">${esc(x.name)}</td>
      <td>${lbl[x.type]} ${esc(x.cat)} ${x.seats ? "· " + x.seats : ""}</td>
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
function buildSummarySheet(wb, s, kind, refDate) {
  const id = lang === "id";
  const ws = wb.addWorksheet(id ? "Ringkasan" : "Summary");
  ws.addRows([
    [id ? "LAPORAN KEUANGAN" : "FINANCIAL REPORT", D().event],
    [t("evDate"), D().date || "-"],
    [id ? "Periode" : "Period", periodLabel(refDate, kind)],
    [id ? "Dicetak" : "Generated", new Date().toLocaleString()],
    [id ? "Oleh" : "By", acting().name + " (" + acting().email + ")"],
    [],
    [t("ticket"), s.inTicket],
    [t("don"), s.inDon],
    [id ? "Total pemasukan" : "Total income", s.inTicket + s.inDon],
    [t("exp"), s.exp],
    [t("net"), s.net],
    [],
    [t("wait"), s.pendAmt],
    [],
    [t("quota"), s.quota],
    [id ? "Kursi terjual" : "Seats sold", s.sold],
    [id ? "Kursi dipesan belum lunas" : "Seats reserved unpaid", s.held],
    [id ? "Kursi tersedia" : "Seats available", s.avail],
  ]);
  ws.getColumn(1).width = 28;
  ws.getColumn(1).font = { bold: true };
  ws.getColumn(2).width = 30;
}
function buildSeatsSheet(wb) {
  const id = lang === "id";
  const ws = wb.addWorksheet(id ? "Kursi" : "Seats");
  ws.addRow([
    t("catTitle"), t("price"), t("quota"),
    id ? "Terjual" : "Sold", id ? "Sisa" : "Left", id ? "Nilai" : "Value",
  ]);
  styleHeaderRow(ws, 1);
  D().cats.forEach((c) => {
    const sd = S.tx
      .filter((x) => x.status === "verified" && x.cat === c.n)
      .reduce((a, b) => a + (+b.seats || 0), 0);
    ws.addRow([c.n, c.p, c.q, sd, c.q - sd, sd * c.p]);
  });
  ws.columns.forEach((c) => (c.width = 16));
}
function buildTrendSheet(wb) {
  const id = lang === "id";
  const ws = wb.addWorksheet(id ? "Tren Harian" : "Daily Trend");
  ws.addRow([t("date"), id ? "Pemasukan" : "Income"]);
  styleHeaderRow(ws, 1);
  seriesByDay(30).forEach((x) => ws.addRow([x.k, x.v]));
  ws.columns.forEach((c) => (c.width = 16));
}
function buildRankSheet(wb) {
  const id = lang === "id";
  const ws = wb.addWorksheet(id ? "Peringkat" : "Rankings");
  ws.addRow([t("topBuy")]);
  ws.getCell(1, 1).font = { bold: true, size: 13 };
  ws.addRow([t("name"), t("trans"), t("seatsW"), "Total"]);
  styleHeaderRow(ws, 2);
  byPerson("ticket").forEach((x) => ws.addRow([x.name, x.n, x.seats, x.amount]));
  ws.addRow([]);
  const r2 = ws.rowCount + 1;
  ws.addRow([t("topDon")]);
  ws.getCell(r2, 1).font = { bold: true, size: 13 };
  ws.addRow([t("name"), t("trans"), t("seatsW"), "Total"]);
  styleHeaderRow(ws, r2 + 1);
  byPerson("donation").forEach((x) => ws.addRow([x.name, x.n, x.seats, x.amount]));
  ws.columns.forEach((c) => (c.width = 20));
}
async function exportExcel(kind, refDate) {
  const s = sums(),
    id = lang === "id";
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = "Kas Acara";
  const txRows = S.tx
    .filter((x) => txInPeriod(x, kind, refDate))
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date));
  const wsTx = wb.addWorksheet(id ? "Transaksi" : "Transactions");
  addReportTitle(wsTx, periodLabel(refDate, kind));
  addTxTable(wsTx, txRows, "TabelTransaksi");
  buildSummarySheet(wb, s, kind, refDate);
  buildSeatsSheet(wb);
  buildTrendSheet(wb);
  buildRankSheet(wb);
  await downloadWorkbook(
    wb,
    `${id ? "Laporan" : "Report"}-${(D().event || "Acara").replace(/\s+/g, "-")}-${refDate || today()}.xlsx`,
  );
  await mutate(() => {}, "actExport", D().event + " · " + rp(s.net));
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
  setType,
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
  toggleArchive,
  loadMonitor,
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
  recalcCash,
  resetAll,
  openExportPeriod,
  setExportKind,
  updateExportPreview,
  doExportExcel,
  exportLogs,
  toggleEye,
});
boot();
