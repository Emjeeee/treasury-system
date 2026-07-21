import './style.css'
import * as XLSX from 'xlsx'
import Chart from 'chart.js/auto'

/* ================= penyimpanan ================= */
const SKEY = "ct_shared_v4",
  LKEY = "ct_session_v4";
let S = {
  rev: 0,
  config: {
    event: "Konser Amal 2026",
    date: "",
    cats: [
      { n: "Reguler", p: 150000, q: 300 },
      { n: "VIP", p: 400000, q: 80 },
    ],
    banks: ["BCA", "Livin Mandiri", "BRI", "Tunai"],
    tpl: [
      { f: "date", h: "Tanggal" },
      { f: "type", h: "Jenis" },
      { f: "name", h: "Nama" },
      { f: "phone", h: "WhatsApp" },
      { f: "cat", h: "Kategori" },
      { f: "seats", h: "Kursi" },
      { f: "amount", h: "Nominal" },
      { f: "bank", h: "Rekening" },
      { f: "status", h: "Status" },
      { f: "note", h: "Catatan" },
    ],
  },
  tx: [],
  users: [],
  logs: [],
  resets: [],
};
let me = null,
  imp = null,
  lang = "id";
let screen = "boot",
  tab = "dash",
  filter = "all",
  panel = "seats",
  adminTab = "users",
  authMode = "in";
let chart1 = null,
  pendingImp = [],
  logQ = "";

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

async function sha(s) {
  const b = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode("kas." + s),
  );
  return [...new Uint8Array(b)]
    .map((x) => x.toString(16).padStart(2, "0"))
    .join("");
}
async function pull() {
  try {
    const r = await window.storage.get(SKEY, true);
    if (r && r.value) {
      const v = JSON.parse(r.value);
      if (v.rev !== S.rev) {
        S = Object.assign(S, v);
        return true;
      }
    }
  } catch (e) {}
  return false;
}
async function push() {
  S.rev = (S.rev || 0) + 1;
  try {
    await window.storage.set(SKEY, JSON.stringify(S), true);
  } catch (e) {
    toast("⚠ " + t("saveErr"));
  }
}
async function mutate(fn, logAct, logDet) {
  try {
    const r = await window.storage.get(SKEY, true);
    if (r && r.value) S = Object.assign(S, JSON.parse(r.value));
  } catch (e) {}
  fn();
  if (logAct) {
    S.logs.unshift({
      ts: now(),
      user: acting().name,
      email: acting().email,
      role: acting().role,
      asAdmin: imp ? me.email : "",
      act: logAct,
      det: logDet || "",
    });
    S.logs = S.logs.slice(0, 600);
  }
  await push();
}
async function saveSession() {
  try {
    await window.storage.set(
      LKEY,
      JSON.stringify({
        email: me ? me.email : "",
        imp: imp ? imp.email : "",
        lang,
      }),
      false,
    );
  } catch (e) {}
}
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
    status: "Status uang",
    stW: "Menunggu cek mutasi",
    stV: "Sudah masuk rekening",
    save: "Simpan transaksi",
    del: "Hapus",
    saved: "Transaksi tersimpan",
    deleted: "Transaksi dihapus",
    verified: "Ditandai sudah masuk rekening",
    needAmt: "Nominal belum diisi",
    event: "Acara",
    evName: "Nama acara",
    evDate: "Tanggal acara",
    bankList: "Rekening / metode pembayaran (pisahkan dengan koma)",
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
    welcome: "Masuk ke Kas Acara",
    welcomeSub: "Catat pemasukan, donasi, dan pengeluaran acara.",
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
    status: "Payment status",
    stW: "Awaiting bank check",
    stV: "Received in account",
    save: "Save transaction",
    del: "Delete",
    saved: "Transaction saved",
    deleted: "Transaction deleted",
    verified: "Marked as received",
    needAmt: "Amount is required",
    event: "Event",
    evName: "Event name",
    evDate: "Event date",
    bankList: "Accounts / payment methods (comma separated)",
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
    welcome: "Sign in to Event Treasury",
    welcomeSub: "Track event income, donations and expenses.",
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
  },
};
const t = (k, v = {}) =>
  String(L[lang][k] || k).replace(/\{(\w+)\}/g, (_, x) => v[x]);
const FL = {
  date: { id: "Tanggal", en: "Date" },
  type: { id: "Jenis", en: "Type" },
  name: { id: "Nama", en: "Name" },
  phone: { id: "WhatsApp", en: "WhatsApp" },
  cat: { id: "Kategori", en: "Category" },
  seats: { id: "Kursi", en: "Seats" },
  amount: { id: "Nominal", en: "Amount" },
  bank: { id: "Rekening", en: "Account" },
  status: { id: "Status", en: "Status" },
  note: { id: "Catatan", en: "Note" },
};

/* ================= hitung ================= */
function sums() {
  const v = S.tx.filter((x) => x.status === "verified");
  const inTicket = v
    .filter((x) => x.type === "ticket")
    .reduce((a, b) => a + b.amount, 0);
  const inDon = v
    .filter((x) => x.type === "donation")
    .reduce((a, b) => a + b.amount, 0);
  const exp = v
    .filter((x) => x.type === "expense")
    .reduce((a, b) => a + b.amount, 0);
  const pd = S.tx.filter((x) => x.status === "pending" && x.type !== "expense");
  const quota = D().cats.reduce((a, c) => a + (+c.q || 0), 0);
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
function byPerson(ty) {
  const m = {};
  S.tx
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
async function boot() {
  try {
    const r = await window.storage.get(LKEY, false);
    if (r && r.value) {
      const v = JSON.parse(r.value);
      lang = v.lang || "id";
      await pull();
      me = S.users.find((u) => u.email === v.email && u.active) || null;
      if (me && v.imp) imp = S.users.find((u) => u.email === v.imp) || null;
    }
  } catch (e) {
    await pull();
  }
  if (!S.users.length) {
    screen = "setup";
  } else if (!me) {
    screen = "auth";
  } else {
    screen = "app";
  }
  render();
  setInterval(async () => {
    if (screen !== "app") return;
    const ch = await pull();
    if (ch) {
      if (me) {
        const u = S.users.find((x) => x.email === me.email);
        if (!u || !u.active) {
          await doSignOut(true);
          return;
        }
        me = u;
        if (imp) imp = S.users.find((x) => x.email === imp.email) || null;
      }
      render();
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
  else {
    r.innerHTML = vApp();
    if (tab === "dash") drawChart();
  }
}
function langSeg() {
  return `<div class="seg">${["id", "en"]
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

/* ================= layar setup & auth ================= */
function vSetup() {
  return `<div class="auth"><div class="authbox">
  <div class="rowsp" style="margin-bottom:14px"><div style="display:flex;gap:10px;align-items:center">
    <div class="mark">K</div><h1 style="font-size:19px">${t("setupT")}</h1></div>${langSeg()}</div>
  <div class="card" style="padding:20px">
    <p class="hint" style="margin:0 0 14px">${t("setupSub")}</p>
    <div class="field"><label>${t("fullName")}</label><input id="a_name"></div>
    <div class="field"><label>${t("email")}</label><input id="a_email" type="email" autocomplete="username"></div>
    <div class="field"><label>${t("pass")}</label><input id="a_pass" type="password" autocomplete="new-password"></div>
    <div class="field"><label>${t("recovNew")}</label><input id="a_rec" value="${uid().toUpperCase()}">
      <div class="hint" style="margin-top:4px">${t("recovHint")}</div></div>
    <button class="btn wide" onclick="createFirst()">${t("createAdmin")}</button>
  </div></div></div>`;
}
async function createFirst() {
  const n = val("a_name"),
    e = val("a_email").toLowerCase(),
    p = val("a_pass"),
    rc = val("a_rec");
  if (!n || !e || !p || !rc) return toast(t("fillAll"));
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
  };
  await mutate(
    () => {
      S.users.push(u);
    },
    "actSignup",
    "admin " + e,
  );
  me = u;
  imp = null;
  await saveSession();
  screen = "app";
  render();
  toast(t("hello", { n }));
}
const val = (i) => (document.getElementById(i)?.value || "").trim();

function vAuth() {
  const f = authMode;
  return `<div class="auth"><div class="authbox">
  <div class="rowsp" style="margin-bottom:14px"><div style="display:flex;gap:10px;align-items:center">
    <div class="mark">K</div><div><h1 style="font-size:19px">${f === "forgot" ? t("forgotT") : t("welcome")}</h1>
    <div class="hint">${f === "forgot" ? t("forgotSub") : t("welcomeSub")}</div></div></div>${langSeg()}</div>
  <div class="card" style="padding:20px">
  ${
    f === "in"
      ? `
    <div class="field"><label>${t("email")}</label><input id="i_email" type="email" autocomplete="username"></div>
    <div class="field"><label>${t("pass")}</label><input id="i_pass" type="password" autocomplete="current-password"
      onkeydown="if(event.key==='Enter')doSignIn()"></div>
    <button class="btn wide" onclick="doSignIn()">${t("signIn")}</button>
    <div class="divider">atau</div>
    <button class="gbtn" onclick="googleFlow()">${gIcon()} ${t("google")}</button>
    <div class="rowsp" style="margin-top:14px">
      <button class="btn ghost sm" onclick="authMode='forgot';render()">${t("forgot")}</button>
      <button class="btn ghost sm" onclick="authMode='up';render()">${t("signUp")}</button></div>`
      : f === "up"
        ? `
    <div class="field"><label>${t("fullName")}</label><input id="u_name"></div>
    <div class="field"><label>${t("email")}</label><input id="u_email" type="email" autocomplete="username"></div>
    <div class="field"><label>${t("pass")}</label><input id="u_pass" type="password" autocomplete="new-password"></div>
    <div class="field"><label>${t("pass2")}</label><input id="u_pass2" type="password" autocomplete="new-password"></div>
    <div class="field"><label>${t("recovNew")}</label><input id="u_rec" value="${uid().toUpperCase()}">
      <div class="hint" style="margin-top:4px">${t("recovHint")}</div></div>
    <button class="btn wide" onclick="doSignUp()">${t("signUp")}</button>
    <div class="divider">atau</div>
    <button class="gbtn" onclick="googleFlow()">${gIcon()} ${t("google")}</button>
    <button class="btn ghost sm wide" style="margin-top:14px" onclick="authMode='in';render()">${t("haveAcc")}</button>`
        : `
    <div class="field"><label>${t("email")}</label><input id="r_email" type="email"></div>
    <div class="field"><label>${t("recov")}</label><input id="r_rec"></div>
    <div class="field"><label>${t("newPass")}</label><input id="r_pass" type="password"></div>
    <button class="btn wide" onclick="doForgot()">${t("resetPass")}</button>
    <p class="hint" style="margin:12px 0 0">${t("reqAdmin")}</p>
    <button class="btn ghost sm wide" style="margin-top:10px" onclick="authMode='in';render()">${t("signIn")}</button>`
  }
  </div></div></div>`;
}
function gIcon() {
  return `<svg width="18" height="18" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.5 0 6.6 1.2 9 3.6l6.7-6.7C35.6 2.6 30.1 0 24 0 14.6 0 6.5 5.4 2.6 13.2l7.8 6.1C12.3 13.2 17.6 9.5 24 9.5z"/><path fill="#4285F4" d="M46.5 24.5c0-1.6-.1-3.2-.4-4.7H24v9h12.7c-.5 3-2.2 5.5-4.7 7.2l7.3 5.7c4.3-4 6.8-9.9 6.8-17.2z"/><path fill="#FBBC05" d="M10.4 28.7c-.5-1.5-.8-3-.8-4.7s.3-3.2.8-4.7l-7.8-6.1C.9 16.4 0 20.1 0 24s.9 7.6 2.6 10.8l7.8-6.1z"/><path fill="#34A853" d="M24 48c6.5 0 11.9-2.1 15.9-5.8l-7.3-5.7c-2 1.4-4.7 2.3-8.6 2.3-6.4 0-11.7-3.7-13.6-8.8l-7.8 6.1C6.5 42.6 14.6 48 24 48z"/></svg>`;
}

async function doSignIn() {
  const e = val("i_email").toLowerCase(),
    p = val("i_pass");
  await pull();
  const u = S.users.find((x) => x.email === e && x.active);
  if (!u || u.pass !== (await sha(e + p))) return toast(t("badLogin"));
  me = u;
  imp = null;
  await mutate(
    () => {
      const x = S.users.find((v) => v.email === e);
      x.last = now();
    },
    "actLogin",
    "",
  );
  await saveSession();
  screen = "app";
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
  await pull();
  if (S.users.some((x) => x.email === e)) return toast(t("emailUsed"));
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
  };
  me = u;
  imp = null;
  await mutate(
    () => {
      S.users.push(u);
    },
    "actSignup",
    e,
  );
  await saveSession();
  screen = "app";
  render();
  toast(t("hello", { n }));
}
function googleFlow() {
  sheet(`<div class="rowsp" style="margin-bottom:10px"><h2 style="font-size:19px">${t("gTitle")}</h2>
    <button class="btn ghost sm" onclick="closeSheet()">${t("close")}</button></div>
    <p class="hint" style="margin:0 0 12px">${t("gSub")}</p>
    <div class="field"><label>${t("email")}</label><input id="g_email" type="email" placeholder="nama@gmail.com"></div>
    <div class="field"><label>${t("fullName")}</label><input id="g_name"></div>
    <button class="btn wide" onclick="doGoogle()">${t("gGo")}</button>`);
}
async function doGoogle() {
  const e = val("g_email").toLowerCase(),
    n = val("g_name") || e.split("@")[0];
  if (!e) return toast(t("fillAll"));
  await pull();
  let u = S.users.find((x) => x.email === e);
  if (u) {
    if (!u.active) return toast(t("badLogin"));
    me = u;
    await mutate(
      () => {
        S.users.find((v) => v.email === e).last = now();
      },
      "actLogin",
      "Google",
    );
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
    };
    me = u;
    await mutate(
      () => {
        S.users.push(u);
      },
      "actSignup",
      e + " · Google",
    );
  }
  imp = null;
  closeSheet();
  await saveSession();
  screen = "app";
  render();
  toast(t("hello", { n: me.name }));
}
async function doForgot() {
  const e = val("r_email").toLowerCase(),
    rc = val("r_rec").toUpperCase(),
    p = val("r_pass");
  if (p.length < 8) return toast(t("passShort"));
  await pull();
  const u = S.users.find((x) => x.email === e);
  if (!u || !u.rec || u.rec !== (await sha(e + rc)))
    return toast(t("badRecov"));
  const h = await sha(e + p);
  await mutate(() => {
    S.users.find((v) => v.email === e).pass = h;
  }, null);
  toast(t("passReset"));
  authMode = "in";
  render();
}
async function doSignOut(silent) {
  if (!silent) await mutate(() => {}, "actLogout", "");
  me = null;
  imp = null;
  await saveSession();
  screen = "auth";
  authMode = "in";
  render();
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
    <div class="mark">K</div>
    <div style="flex:1;min-width:0"><h1 style="font-size:17px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(D().event || "Konser")}</h1>
      <div class="sub">${(D().date ? new Date(D().date + "T00:00:00").toLocaleDateString(lang === "id" ? "id-ID" : "en-GB", { day: "numeric", month: "short", year: "numeric" }) + " · " : "") + s.sold + " " + t("seatsSold")}</div></div>
    ${langSeg()}
    <button class="btn ghost sm" onclick="exportExcel()">${t("xls")}</button>
    <div class="avatar" onclick="openMe()" title="${esc(a.name)}">${esc(a.name.slice(0, 2).toUpperCase())}</div>
  </div></header>
  <div class="wrap">${{ dash: vDash, tx: vTx, board: vBoard, admin: vAdmin }[tab]()}</div>
  ${tab === "tx" || tab === "dash" ? `<button class="btn fab" onclick="openTx()">${t("add")}</button>` : ""}
  <nav>${tabs.map(([k, l]) => `<button class="${tab === k ? "on" : ""}" onclick="go('${k}')"><span class="dot"></span>${l}</button>`).join("")}</nav>`;
}
function go(k) {
  tab = k;
  render();
}
function openMe() {
  const a = acting();
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${esc(a.name)}</h2>
    <button class="btn ghost sm" onclick="closeSheet()">${t("close")}</button></div>
    <div class="hint" style="margin-bottom:6px">${esc(a.email)}</div>
    <span class="tag ${a.role === "admin" ? "t-adm" : "t-tic"}">${a.role === "admin" ? t("admins") : t("treas")}</span>
    ${
      imp
        ? `<p class="hint" style="margin-top:12px">${t("byAdmin", { n: esc(me.name) })}</p>
      <button class="btn wide" style="margin-top:8px" onclick="stopImp()">${t("backAdmin")}</button>`
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
            `<button class="${panel === p ? "on" : ""}" onclick="panel='${p}';render()">${l}</button>`,
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
  chart1 && chart1.destroy();
  chart1 = new Chart(el, {
    type: "bar",
    data: {
      labels: s.map((x) => x.label),
      datasets: [
        {
          data: s.map((x) => x.v),
          backgroundColor: "#2F6FEB",
          hoverBackgroundColor: "#2258C4",
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
          ticks: { color: "#6B6B66", font: { size: 12 } },
        },
        y: {
          grid: { color: "#F0EFEC" },
          border: { display: false },
          ticks: {
            color: "#6B6B66",
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
function vTx() {
  const f = S.tx
    .filter(
      (x) =>
        filter === "all" ||
        (filter === "pending" ? x.status === "pending" : x.type === filter),
    )
    .sort((a, b) => (b.date + b.id).localeCompare(a.date + a.id));
  return `<div style="padding:12px 0 110px">
  <div class="rowsp" style="margin-bottom:11px;flex-wrap:wrap">
    <div class="chips">${[
      ["all", t("all")],
      ["pending", t("unchecked")],
      ["ticket", t("seatsW")],
      ["donation", t("donW")],
      ["expense", t("expW")],
    ]
      .map(
        ([k, l]) =>
          `<button class="chip ${filter === k ? "on" : ""}" onclick="filter='${k}';render()">${l}</button>`,
      )
      .join("")}</div>
    <button class="btn ghost sm" onclick="openImport()">${t("imp")}</button></div>
  <div class="card" style="padding:2px 10px">${
    f.length
      ? `<div class="scroll" style="max-height:calc(100dvh - 210px)"><table><thead><tr>
    <th>${t("date")}</th><th>${t("name")}</th><th>${t("detail")}</th><th style="text-align:right">${t("amount")}</th><th></th></tr></thead><tbody>
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
        <div class="hint">${esc(x.bank || "-")}${x.note ? " · " + esc(x.note) : ""}</div></td>
      <td class="mono ${x.type === "expense" ? "neg" : ""}" style="text-align:right;white-space:nowrap;font-weight:600">${x.type === "expense" ? "−" : ""}${rp(x.amount)}
        <div style="margin-top:3px">${x.status === "verified" ? `<span class="tag t-ok">${t("inRek")}</span>` : `<span class="tag t-wait">${t("stW")}</span>`}</div></td>
      <td style="text-align:right;white-space:nowrap">${x.status === "pending" ? `<button class="btn ok sm" onclick="verify('${x.id}')">${t("inRek")}</button> ` : ""}
        <button class="btn ghost sm" onclick="openTx('${x.id}')">${t("edit")}</button></td></tr>`,
      )
      .join("")}
    </tbody></table></div>`
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
        bank: D().banks[0] || "",
        note: "",
        status: "pending",
      };
  sheet(`<div class="rowsp" style="margin-bottom:14px"><h2 style="font-size:20px">${id ? t("editTx") : t("newTx")}</h2>
      <button class="btn ghost sm" onclick="closeSheet()">${t("close")}</button></div>
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
      <div class="field"><label id="l_bank">${t("bankTo")}</label><select id="f_bank">${D()
        .banks.map(
          (b) => `<option ${b === x.bank ? "selected" : ""}>${esc(b)}</option>`,
        )
        .join("")}</select></div>
      <div class="field"><label id="l_name">${t("nameBuy")}</label><input id="f_name" value="${esc(x.name)}"></div>
      <div class="field" id="w_phone"><label>${t("wa")}</label><input id="f_phone" inputmode="tel" value="${esc(x.phone)}" placeholder="0812..."></div>
      <div class="field" id="w_cat"><label id="l_cat">${t("catSeat")}</label><select id="f_cat" onchange="recalc()">${D()
        .cats.map(
          (c) =>
            `<option ${c.n === x.cat ? "selected" : ""}>${esc(c.n)}</option>`,
        )
        .join("")}</select></div>
      <div class="field" id="w_seats"><label>${t("nSeat")}</label><input type="number" inputmode="numeric" min="0" id="f_seats" value="${x.seats || 0}" oninput="recalc()"></div>
    </div>
    <div class="field" id="w_expcat" style="display:none"><label>${t("purpose")}</label><input id="f_expcat" value="${x.type === "expense" ? esc(x.cat) : ""}"></div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field"><label id="l_amount">${t("amtIn")}</label><input type="number" inputmode="numeric" id="f_amount" value="${x.amount}" oninput="prev()">
        <div class="hint mono" id="prev" style="margin-top:4px">${rp(x.amount)}</div></div>
      <div class="field"><label>${t("status")}</label><select id="f_status">
        <option value="pending" ${x.status === "pending" ? "selected" : ""}>${t("stW")}</option>
        <option value="verified" ${x.status === "verified" ? "selected" : ""}>${t("stV")}</option></select></div></div>
    <div class="field"><label>${t("note")}</label><textarea id="f_note" rows="2">${esc(x.note)}</textarea></div>
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
async function saveTx() {
  const g = (i) => document.getElementById("f_" + i).value,
    ty = g("type"),
    id = g("id");
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
    note: g("note"),
    status: g("status"),
    by: acting().name,
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
  const a = byPerson("ticket").slice(0, 15),
    b = byPerson("donation").slice(0, 15);
  const row = (
    x,
    i,
    u,
  ) => `<div class="lb" style="animation-delay:${i * 35}ms"><div class="rank ${i === 0 ? "g1" : ""}">${i + 1}</div>
    <div style="flex:1;min-width:0"><div style="font-weight:600">${esc(x.name)}</div>
    <div class="hint">${x.n} ${t("trans")} · ${x.seats} ${t("seatsW").toLowerCase()}${u}</div></div>
    <div class="mono" style="font-weight:700">${rp(x.amount)}</div></div>`;
  return `<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(330px,1fr));padding:12px 0 110px">
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("topBuy")}</h2>
      ${a.map((x, i) => row(x, i, "")).join("") || `<div class="empty">${t("noneYet")}</div>`}</div>
    <div class="card"><h2 style="font-size:17px;margin-bottom:4px">${t("topDon")}</h2>
      ${b.map((x, i) => row(x, i, " " + t("free"))).join("") || `<div class="empty">${t("noneYet")}</div>`}</div></div>`;
}

/* ================= admin ================= */
function vAdmin() {
  if (!isAdmin()) return `<div class="empty">${t("onlyAdmin")}</div>`;
  return `<div style="padding:12px 0 110px">
    <div class="seg" style="margin-bottom:12px">${[
      ["users", t("users")],
      ["logs", t("logs")],
      ["set", t("set")],
    ]
      .map(
        ([k, l]) =>
          `<button class="${adminTab === k ? "on" : ""}" onclick="adminTab='${k}';render()">${l}</button>`,
      )
      .join("")}</div>
    ${{ users: vUsers, logs: vLogs, set: vSet }[adminTab]()}</div>`;
}
function vUsers() {
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px">
    <h2 style="font-size:17px">${t("users")}</h2><button class="btn sm" onclick="openUser()">${t("addUser")}</button></div>
  <div style="overflow-x:auto"><table><thead><tr><th>${t("name")}</th><th>${t("role")}</th><th>${t("status")}</th>
    <th>${t("lastIn")}</th><th></th></tr></thead><tbody>
    ${S.users
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
      .join("")}
  </tbody></table></div></div>`;
}
function openUser(email) {
  const u = email
    ? S.users.find((x) => x.email === email)
    : {
        name: "",
        email: "",
        role: "treasurer",
        active: true,
        provider: "password",
      };
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:19px">${email ? t("edit") : t("addUser")}</h2>
    <button class="btn ghost sm" onclick="closeSheet()">${t("close")}</button></div>
    <input type="hidden" id="v_old" value="${esc(email || "")}">
    <div class="field"><label>${t("fullName")}</label><input id="v_name" value="${esc(u.name)}"></div>
    <div class="field"><label>${t("email")}</label><input id="v_email" type="email" value="${esc(u.email)}" ${email ? "disabled" : ""}></div>
    <div class="grid" style="grid-template-columns:1fr 1fr">
      <div class="field"><label>${t("role")}</label><select id="v_role">
        <option value="treasurer" ${u.role === "treasurer" ? "selected" : ""}>${t("treas")}</option>
        <option value="admin" ${u.role === "admin" ? "selected" : ""}>${t("admins")}</option></select></div>
      <div class="field"><label>${t("status")}</label><select id="v_active">
        <option value="1" ${u.active ? "selected" : ""}>${t("active")}</option>
        <option value="0" ${!u.active ? "selected" : ""}>${t("inactive")}</option></select></div></div>
    <div class="field"><label>${email ? t("newPass") : t("pass")}</label><input id="v_pass" type="password" placeholder="${email ? "—" : ""}"></div>
    <div class="rowsp" style="margin-top:12px">
      ${email && email !== me.email ? `<button class="btn danger" onclick="delUser('${esc(email)}')">${t("del")}</button>` : "<span></span>"}
      <button class="btn" onclick="saveUser()">${t("save")}</button></div>`);
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
  await pull();
  if (!old && S.users.some((x) => x.email === e)) return toast(t("emailUsed"));
  if (!old && !p) return toast(t("fillAll"));
  const admins = S.users.filter(
    (x) => x.role === "admin" && x.active && x.email !== old,
  ).length;
  if (old && !(role === "admin" && act) && admins === 0)
    return toast(t("lastAdmin"));
  const hash = p ? await sha(e + p) : null;
  await mutate(
    () => {
      if (old) {
        const u = S.users.find((x) => x.email === old);
        u.name = n;
        u.role = role;
        u.active = act;
        if (hash) u.pass = hash;
      } else
        S.users.push({
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
        });
    },
    "actUser",
    `${old ? t("userSaved") : t("userAdded")}: ${n} (${role === "admin" ? t("admins") : t("treas")})`,
  );
  if (me.email === old) me = S.users.find((x) => x.email === old);
  closeSheet();
  toast(old ? t("userSaved") : t("userAdded"));
  render();
}
async function delUser(email) {
  if (!confirm(t("confirmUser"))) return;
  await pull();
  const u = S.users.find((x) => x.email === email);
  if (
    S.users.filter((x) => x.role === "admin" && x.active).length <= 1 &&
    u.role === "admin"
  )
    return toast(t("lastAdmin"));
  await mutate(
    () => {
      S.users = S.users.filter((x) => x.email !== email);
    },
    "actUser",
    t("userDel") + ": " + u.name,
  );
  closeSheet();
  toast(t("userDel"));
  render();
}
async function startImp(email) {
  imp = S.users.find((x) => x.email === email);
  tab = "dash";
  await mutate(() => {}, "actImp", imp.name + " (" + imp.email + ")");
  await saveSession();
  render();
  toast(t("impBanner", { n: imp.name }));
}

/* ================= log aktivitas ================= */
function vLogs() {
  const q = logQ.toLowerCase();
  const rows = S.logs.filter(
    (l) =>
      !q ||
      [l.user, l.email, t(l.act), l.det].join(" ").toLowerCase().includes(q),
  );
  return `<div class="card"><div class="rowsp" style="margin-bottom:8px;flex-wrap:wrap;gap:8px">
    <h2 style="font-size:17px">${t("logs")}</h2>
    <div style="display:flex;gap:8px;flex:1;justify-content:flex-end;min-width:220px">
      <input style="max-width:320px" placeholder="${t("searchLog")}" value="${esc(logQ)}"
        oninput="logQ=this.value;clearTimeout(window._lq);window._lq=setTimeout(render,250)">
      <button class="btn ghost sm" onclick="exportLogs()">${t("exportLog")}</button></div></div>
  ${
    rows.length
      ? `<div class="scroll" style="max-height:calc(100dvh - 260px)"><table><thead><tr>
    <th>${t("time")}</th><th>${t("user")}</th><th>${t("action")}</th><th>${t("info")}</th></tr></thead><tbody>
    ${rows
      .map(
        (l) => `<tr>
      <td class="hint mono" style="white-space:nowrap">${new Date(l.ts).toLocaleString(lang === "id" ? "id-ID" : "en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</td>
      <td><div style="font-weight:600">${esc(l.user)}</div><div class="hint">${esc(l.email)}${l.asAdmin ? " · " + t("byAdmin", { n: esc(l.asAdmin) }) : ""}</div></td>
      <td><span class="tag ${/Delete|Hapus/.test(t(l.act)) ? "t-exp" : /Verif/.test(t(l.act)) ? "t-ok" : "t-tic"}">${t(l.act)}</span></td>
      <td class="hint">${esc(l.det)}</td></tr>`,
      )
      .join("")}</tbody></table></div>`
      : `<div class="empty">${t("noLog")}</div>`
  }</div>`;
}
function exportLogs() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [t("time"), t("user"), t("email"), t("role"), t("action"), t("info")],
      ...S.logs.map((l) => [
        new Date(l.ts).toLocaleString(),
        l.user,
        l.email,
        l.role,
        t(l.act),
        l.det,
      ]),
    ]),
    "Log",
  );
  XLSX.writeFile(wb, `Activity-Log-${today()}.xlsx`);
  toast(t("exportLog"));
}

/* ================= pengaturan (admin) ================= */
function vSet() {
  return `<div class="grid" style="gap:10px">
  <div class="card"><h2 style="font-size:17px;margin-bottom:10px">${t("event")}</h2>
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))">
      <div class="field"><label>${t("evName")}</label><input id="s_ev" value="${esc(D().event)}"></div>
      <div class="field"><label>${t("evDate")}</label><input type="date" id="s_date" value="${D().date || ""}"></div></div>
    <div class="field"><label>${t("bankList")}</label><input id="s_banks" value="${esc(D().banks.join(", "))}"></div>
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
    <div class="card"><h2 style="font-size:17px">${t("tplTitle")}</h2>
      <p class="hint" style="margin:4px 0 8px">${t("tplP")}</p>
      <div style="overflow-x:auto"><table><thead><tr><th>${t("field")}</th><th>${t("colName")}</th></tr></thead><tbody>
      ${D()
        .tpl.map(
          (c, i) => `<tr><td style="font-weight:600">${FL[c.f][lang]}</td>
        <td><input value="${esc(c.h)}" onchange="editTpl(${i},this.value)"></td></tr>`,
        )
        .join("")}
      </tbody></table></div>
      <div style="display:flex;gap:8px;margin-top:10px;flex-wrap:wrap">
        <button class="btn ghost sm" onclick="dlTemplate()">${t("dlTpl")}</button>
        <button class="btn sm" onclick="openImport()">${t("imp")}</button></div></div></div>
  <div class="card"><h2 style="font-size:17px">${t("dataT")}</h2>
    <p class="hint" style="margin:4px 0 10px">${t("dataP")}</p>
    <div><button class="btn danger" onclick="resetAll()">${t("clearAll")}</button></div></div></div>`;
}
async function saveSet() {
  const ev = val("s_ev"),
    dt = val("s_date"),
    bk = val("s_banks")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
  await mutate(
    () => {
      S.config.event = ev;
      S.config.date = dt;
      S.config.banks = bk;
    },
    "actSet",
    ev + " · " + bk.join(", "),
  );
  toast(t("setSaved"));
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

/* ================= impor ================= */
function dlTemplate() {
  const hdr = D().tpl.map((c) => c.h),
    id = lang === "id";
  const ex = [
    [
      today(),
      id ? "Kursi" : "Seat",
      "Budi Santoso",
      "081234567890",
      D().cats[0]?.n || "Reguler",
      2,
      (D().cats[0]?.p || 0) * 2,
      D().banks[0] || "BCA",
      id ? "Masuk" : "Received",
      "Transfer 08.41",
    ],
    [
      today(),
      id ? "Donasi" : "Donation",
      "PT Nada Jaya",
      "",
      D().cats[0]?.n || "Reguler",
      10,
      25000000,
      D().banks[0] || "BCA",
      id ? "Menunggu" : "Pending",
      "",
    ],
    [
      today(),
      id ? "Pengeluaran" : "Expense",
      "Sound System",
      "",
      "",
      0,
      7500000,
      D().banks[0] || "BCA",
      id ? "Masuk" : "Received",
      "DP 50%",
    ],
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([hdr, ...ex]),
    "Template",
  );
  XLSX.writeFile(wb, "Template-Import.xlsx");
  toast(t("dlTpl"));
}
function openImport() {
  sheet(`<div class="rowsp" style="margin-bottom:12px"><h2 style="font-size:20px">${t("imp")}</h2>
    <button class="btn ghost sm" onclick="closeSheet()">${t("close")}</button></div>
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
function parseRows(rows) {
  const body = document.getElementById("impBody");
  if (!rows || rows.length < 2)
    return (body.innerHTML = `<div class="empty">${t("noCol")}</div>`);
  const head = rows[0].map((h) => String(h).trim().toLowerCase()),
    idx = {};
  D().tpl.forEach((c) => {
    const i = head.indexOf(String(c.h).trim().toLowerCase());
    if (i >= 0) idx[c.f] = i;
  });
  if (idx.amount === undefined && idx.name === undefined)
    return (body.innerHTML = `<div class="empty">${t("noCol")}</div>`);
  const get = (r, f) => (idx[f] === undefined ? "" : r[idx[f]]);
  pendingImp = [];
  let bad = 0;
  rows.slice(1).forEach((r) => {
    if (!r || !r.length || r.every((c) => String(c).trim() === "")) return;
    const amt = normNum(get(r, "amount")),
      dt = normDate(get(r, "date")) || today();
    if (!amt) {
      bad++;
      return;
    }
    const ty = normType(get(r, "type"));
    const st = /masuk|verif|lunas|received|paid|ok/i.test(
      String(get(r, "status")),
    )
      ? "verified"
      : "pending";
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
      bank: String(get(r, "bank") || D().banks[0] || ""),
      note: String(get(r, "note") || ""),
      status: st,
      by: acting().name,
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
async function exportExcel() {
  const s = sums(),
    wb = XLSX.utils.book_new(),
    id = lang === "id";
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [id ? "LAPORAN KEUANGAN" : "FINANCIAL REPORT", D().event],
      [t("evDate"), D().date || "-"],
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
    ]),
    id ? "Ringkasan" : "Summary",
  );
  const lbl = { ticket: t("tBuy"), donation: t("tDon"), expense: t("tExp") };
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [...D().tpl.map((c) => c.h), id ? "Dicatat oleh" : "Recorded by"],
      ...S.tx
        .slice()
        .sort((a, b) => a.date.localeCompare(b.date))
        .map((x) => [
          ...D().tpl.map((c) =>
            c.f === "type"
              ? lbl[x.type]
              : c.f === "status"
                ? x.status === "verified"
                  ? t("inRek")
                  : t("stW")
                : c.f === "amount"
                  ? x.type === "expense"
                    ? -x.amount
                    : x.amount
                  : x[c.f],
          ),
          x.by || "",
        ]),
    ]),
    id ? "Transaksi" : "Transactions",
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [
        t("catTitle"),
        t("price"),
        t("quota"),
        id ? "Terjual" : "Sold",
        id ? "Sisa" : "Left",
        id ? "Nilai" : "Value",
      ],
      ...D().cats.map((c) => {
        const sd = S.tx
          .filter((x) => x.status === "verified" && x.cat === c.n)
          .reduce((a, b) => a + (+b.seats || 0), 0);
        return [c.n, c.p, c.q, sd, c.q - sd, sd * c.p];
      }),
    ]),
    id ? "Kursi" : "Seats",
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [t("date"), id ? "Pemasukan" : "Income"],
      ...seriesByDay(30).map((x) => [x.k, x.v]),
    ]),
    id ? "Tren Harian" : "Daily Trend",
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [t("topBuy")],
      [t("name"), t("trans"), t("seatsW"), "Total"],
      ...byPerson("ticket").map((x) => [x.name, x.n, x.seats, x.amount]),
      [],
      [t("topDon")],
      [t("name"), t("trans"), t("seatsW"), "Total"],
      ...byPerson("donation").map((x) => [x.name, x.n, x.seats, x.amount]),
    ]),
    id ? "Peringkat" : "Rankings",
  );
  XLSX.writeFile(
    wb,
    `${id ? "Laporan" : "Report"}-${(D().event || "Konser").replace(/\s+/g, "-")}-${today()}.xlsx`,
  );
  await mutate(() => {}, "actExport", D().event + " · " + rp(s.net));
  toast(t("xls"));
}
boot();
