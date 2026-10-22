/**
 * ==========================================================================
 * KONFIGURASI DAN STATE GLOBAL APLIKASI
 * ==========================================================================
 */

// Kunci penyimpanan di localStorage browser
const KEY = "toko-v2";

// State utama aplikasi:
// - master : Daftar semua barang master [{ id, kat, sup, sat, kode, nama }]
// - map    : Pemetaan barcode ke ID barang { [barcode]: id }
// - log    : Riwayat scan barang [{ t (timestamp), bc (barcode), id, qty }]
let S = { master: [], map: {}, log: [] };

// Variabel bantu untuk performa dan interaktivitas:
let byId = {};      // Kamus referensi barang cepat berdasarkan ID: { [id]: barang }
let hay = {};       // Kamus teks pencarian huruf kecil: { [id]: "nama kode" }
let card = null;    // State kartu aktif saat proses scan / pencocokan barang
let handle = null;  // File handle saat terhubung langsung ke file Excel lokal
let stream = null;  // Objek MediaStream dari kamera aktif
let timer = null;   // Timer interval untuk looping deteksi barcode dari video
let last = "";      // Barcode terakhir yang dipindai (untuk debounce)
let lastT = 0;      // Timestamp pemindaian terakhir (mencegah double scan cepat)
let saving = false; // Status apakah sedang proses menyimpan ke file Excel
let dirty = false;  // Penanda apakah ada perubahan data baru saat proses simpan masih berjalan

// Muat data tersimpan dari localStorage browser jika ada
try {
  const x = JSON.parse(localStorage.getItem(KEY) || "null");
  if (x && x.master) S = x;
} catch (e) {}

/**
 * ==========================================================================
 * FUNGSI-FUNGSI HELPER & UTILITAS
 * ==========================================================================
 */

/**
 * Mengambil elemen DOM berdasarkan atribut ID-nya (shortcut document.getElementById).
 * @param {string} id - ID dari elemen HTML yang ingin diambil.
 * @returns {HTMLElement|null} Elemen DOM yang cocok.
 */
const $ = (id) => document.getElementById(id);

/**
 * Membuat elemen HTML baru secara dinamis beserta atribut dan anak-anaknya.
 * @param {string} t - Tag name elemen HTML (misal: "div", "button", "td").
 * @param {Object} [p] - Objek atribut / properti yang akan di-assign ke elemen.
 * @param {...(Node|string)} k - Daftar child node atau teks yang akan dimasukkan ke elemen.
 * @returns {HTMLElement} Elemen HTML yang sudah dibuat dan diisi.
 */
function h(t, p, ...k) {
  const e = document.createElement(t);
  Object.assign(e, p || {});
  k.forEach((x) => e.append(x));
  return e;
}

/**
 * Menampilkan pesan status atau notifikasi ke pengguna pada baris status (#status).
 * @param {string} t - Pesan teks yang ingin ditampilkan.
 * @param {boolean} [warn=false] - Jika true, pesan diberi warna peringatan (merah).
 */
function say(t, warn) {
  const s = $("status");
  s.textContent = t;
  s.className = warn ? "warn" : "";
}

/**
 * Memformat timestamp milidetik menjadi teks tanggal dan waktu format Indonesia.
 * @param {number} d - Timestamp milidetik (misal: Date.now()).
 * @returns {string} String waktu terformat (contoh: "22/10/2026, 10.30.00").
 */
function fmt(d) {
  return new Date(d).toLocaleString("id-ID", {
    dateStyle: "short",
    timeStyle: "medium",
  });
}

/**
 * Mengindeks ulang data master barang ke dictionary `byId` dan `hay`.
 * Berfungsi untuk mempercepat lookup barang by ID dan mempercepat fitur pencarian nama/kode.
 */
function index() {
  byId = {};
  hay = {};
  S.master.forEach((m) => {
    byId[m.id] = m;
    hay[m.id] = (m.nama + " " + m.kode).toLowerCase();
  });
}

/**
 * Mencari data barang master berdasarkan kata kunci pencarian.
 * Mendukung pencarian multi-kata (semua kata kunci harus cocok) dan membatasi hasil maksimal 8 barang.
 * @param {string} q - Kata kunci pencarian nama atau kode barang.
 * @returns {Array<Object>} Daftar barang yang sesuai kriteria pencarian.
 */
function find(q) {
  const t = q.toLowerCase().split(/\s+/).filter(Boolean);
  if (!t.length) return [];
  return S.master
    .filter((m) => t.every((w) => hay[m.id].includes(w)))
    .slice(0, 8);
}

/**
 * Mengembalikan fokus kursor ke kolom input scanner manual / USB (#man).
 * Memudahkan pengguna untuk langsung melakukan scan berikutnya tanpa perlu klik mouse.
 */
function focusMan() {
  setTimeout(() => $("man").focus(), 0);
}

/**
 * ==========================================================================
 * FUNGSI-FUNGSI TAMPILAN (RENDER UI)
 * ==========================================================================
 */

/**
 * Merender informasi status ringkasan master data dan memperbarui tabel riwayat scan.
 * Menampilkan hingga 200 riwayat scan terbaru secara terbalik (terbaru di atas)
 * serta menyediakan tombol hapus pada setiap baris log.
 */
function render() {
  const n = new Set(Object.values(S.map)).size;
  $("stat").textContent = S.master.length
    ? S.master.length +
      " barang di master, " +
      n +
      " sudah punya barcode, " +
      S.log.length +
      " catatan scan"
    : "";
  $("empty").textContent = S.master.length
    ? "Belum ada catatan. Scan barcode pertama."
    : "Impor file master_barang.xlsx dulu untuk memulai.";
  $("empty").style.display = S.log.length ? "none" : "block";

  const b = $("body");
  b.innerHTML = "";
  for (let i = S.log.length - 1, c = 0; i >= 0 && c < 200; i--, c++) {
    const l = S.log[i],
      m = byId[l.id],
      tr = document.createElement("tr");

    [fmt(l.t), l.bc, m ? m.nama : "(barang dihapus)", l.qty].forEach((v) =>
      tr.append(h("td", { textContent: v }))
    );

    const x = h("button", { className: "alt", textContent: "Hapus" });
    x.onclick = () => {
      S.log.splice(i, 1);
      after();
    };
    tr.append(h("td", {}, x));
    b.append(tr);
  }
}

/**
 * Merender kartu interaktif konfirmasi scan (#card).
 * - Jika barcode belum terdaftar (pending): Menampilkan form pencarian barang master untuk menghubungkan barcode,
 *   atau opsi menambahkan barang baru jika belum ada di master.
 * - Jika barcode sudah terdaftar: Menampilkan konfirmasi nama barang, input jumlah (qty),
 *   tombol "Catat", dan tombol "Salah barang" untuk membatalkan relasi barcode.
 */
function renderCard() {
  const el = $("card");
  el.innerHTML = "";
  if (!card) return;

  const box = h("div", { className: "cardbox" });
  const cancel = h("button", {
    className: "alt",
    textContent: "Batal",
    type: "button",
  });
  cancel.onclick = () => {
    card = null;
    renderCard();
    focusMan();
  };

  // Kondisi 1: Barcode baru belum terhubung dengan barang manapun
  if (card.pending) {
    const title = h(
      "p",
      {},
      h("span", { textContent: "Barcode " }),
      h("span", { className: "badge-bc", textContent: card.bc }),
      h("span", { textContent: " belum terdaftar." })
    );
    const sub = h("p", {
      className: "sub",
      style: "margin-bottom:8px;",
      textContent: "Cari dari master Excel atau input sebagai barang baru:",
    });

    const q = h("input", {
      type: "text",
      placeholder: "Ketik nama atau kode barang master...",
    });
    const res = h("div", { className: "res" });

    q.oninput = () => {
      res.innerHTML = "";
      find(q.value).forEach((m) => {
        const b = h(
          "button",
          { className: "alt", type: "button" },
          m.nama,
          h("small", { textContent: "  (" + m.kode + ")" })
        );
        b.onclick = () => pick(m.id);
        res.append(b);
      });
    };

    q.onkeydown = (e) => {
      if (e.key === "Enter") {
        const f = find(q.value)[0];
        if (f) pick(f.id);
      }
    };

    // Wadah formulir inline untuk input barang baru
    const newBox = h("div", { className: "new-box", style: "display:none" });
    const newNameInput = h("input", {
      type: "text",
      placeholder: "Ketik nama barang baru lalu Enter...",
    });
    const saveNewBtn = h("button", {
      type: "button",
      textContent: "Simpan & Hubungkan",
    });
    const cancelNewBtn = h("button", {
      className: "alt",
      type: "button",
      textContent: "Batal",
    });

    const submitNew = () => {
      const val = newNameInput.value.trim();
      if (!val) {
        newNameInput.focus();
        return;
      }
      createNewItem(val);
    };

    saveNewBtn.onclick = submitNew;
    newNameInput.onkeydown = (e) => {
      if (e.key === "Enter") submitNew();
    };

    cancelNewBtn.onclick = () => {
      newBox.style.display = "none";
      addToggleBtn.style.display = "";
      q.focus();
    };

    newBox.append(
      h("p", {
        style: "font-weight:600;margin-bottom:6px;",
        textContent: "Input Barang Baru:",
      }),
      h("div", { className: "row", style: "margin-top:0" }, newNameInput),
      h("div", { className: "row" }, saveNewBtn, cancelNewBtn)
    );

    const addToggleBtn = h("button", {
      className: "alt",
      type: "button",
      textContent: "+ Input Barang Baru",
    });
    addToggleBtn.onclick = () => {
      newBox.style.display = "";
      addToggleBtn.style.display = "none";
      newNameInput.focus();
    };

    box.append(
      title,
      sub,
      h("div", { className: "row", style: "margin-top:0" }, q),
      res,
      newBox,
      h("div", { className: "row" }, addToggleBtn, cancel)
    );
    el.append(box);
    q.focus();
    return;
  }

  // Kondisi 2: Barcode sudah dikenal, tampilkan form konfirmasi nama & jumlah
  const m = byId[card.id];
  const nm = h("input", {
    type: "text",
    value: m.nama,
    "aria-label": "Nama barang",
  });
  const qty = h("input", {
    type: "number",
    value: 1,
    min: 1,
    "aria-label": "Jumlah",
  });
  const ok = h("button", { type: "button", textContent: "Catat" });

  const go = () => commit(nm.value, parseInt(qty.value, 10) || 1);
  ok.onclick = go;
  nm.onkeydown = qty.onkeydown = (e) => {
    if (e.key === "Enter") go();
  };

  const re = h("button", {
    className: "alt",
    type: "button",
    textContent: "Salah barang",
  });
  re.onclick = () => {
    delete S.map[card.bc];
    card = { bc: card.bc, pending: true };
    persist();
    renderCard();
  };

  box.append(
    h("p", {}, h("small", { textContent: "Barcode " + card.bc + " · " + m.kode })),
    h("div", { className: "row", style: "margin-top:0" }, nm, qty),
    h("div", { className: "row" }, ok, re, cancel)
  );
  el.append(box);
  qty.focus();
  qty.select();
}

/**
 * ==========================================================================
 * FUNGSI-FUNGSI LOGIKA & OPERASI DATA
 * ==========================================================================
 */

/**
 * Menyimpan seluruh state data aplikasi (master, pemetaan barcode, log scan) ke localStorage browser.
 * Menampilkan peringatan jika kapasitas penyimpanan browser melebihi batas.
 */
function persist() {
  try {
    localStorage.setItem(KEY, JSON.stringify(S));
  } catch (e) {
    say("Penyimpanan browser penuh. Unduh Excel sekarang.", true);
  }
}

/**
 * Callback umum yang dipanggil setiap kali ada perubahan data:
 * 1. Menyimpan ke localStorage (`persist`).
 * 2. Memperbarui tampilan UI (`render`).
 * 3. Menyimpan sinkronisasi ke file Excel jika terhubung (`saveLinked`).
 */
function after() {
  persist();
  render();
  saveLinked();
}

/**
 * Menghubungkan barcode yang sedang aktif ke barang master tertentu berdasarkan ID barang.
 * @param {number} id - ID barang master yang dipilih.
 */
function pick(id) {
  S.map[card.bc] = id;
  persist();
  card = { bc: card.bc, id };
  renderCard();
}

/**
 * Menambahkan data barang baru secara langsung ke master data (S.master),
 * membuatkan ID auto-increment dan kode unik "BRG[id]", lalu langsung
 * menghubungkannya ke nomor barcode fisik yang sedang aktif.
 * @param {string} name - Nama barang baru yang diinput.
 */
function createNewItem(name) {
  name = (name || "").trim();
  if (!name) return;
  const id = S.master.reduce((a, m) => Math.max(a, m.id), 0) + 1;
  const newItem = {
    id,
    kat: "",
    sup: "",
    sat: "",
    kode: "BRG" + id,
    nama: name,
  };
  S.master.push(newItem);
  index();
  pick(id);
}

/**
 * Fallback dialog prompt jika pengguna memilih input cepat lewat prompt.
 */
function addNew() {
  const n = prompt("Nama barang baru:");
  if (n && n.trim()) createNewItem(n.trim());
}

/**
 * Memberikan feedback visual dan audio saat proses pemindaian berhasil:
 * - Membunyikan suara beep pendek via Web Audio API (jika checkbox #beep aktif).
 * - Memberikan efek animasi flash hijau pada kotak kamera.
 */
function flash() {
  if ($("beep").checked) {
    try {
      const a = new AudioContext(),
        o = a.createOscillator();
      o.frequency.value = 880;
      o.connect(a.destination);
      o.start();
      o.stop(a.currentTime + 0.12);
    } catch (e) {}
  }
  const f = $("fl");
  f.classList.add("on");
  setTimeout(() => f.classList.remove("on"), 60);
}

/**
 * Mencatat item barang hasil scan ke riwayat log scan (`S.log`).
 * @param {string} name - Nama barang (jika ada pembaruan nama).
 * @param {number} qty - Jumlah/kuantitas barang yang masuk.
 */
function commit(name, qty) {
  const m = byId[card.id];
  if (name.trim()) m.nama = name.trim();
  S.log.push({ t: Date.now(), bc: card.bc, id: card.id, qty });
  say("Tercatat: " + m.nama + " × " + qty);
  card = null;
  renderCard();
  flash();
  after();
  focusMan();
}

/**
 * Memproses input kode barcode (baik dari kamera otomatis maupun scanner USB / input teks manual).
 * Menangani debounce 2.5 detik untuk barcode identik agar tidak terinput berulang secara tidak sengaja,
 * serta mode "Langsung tambah 1" (fast mode).
 * @param {string} v - Nilai string barcode yang terbaca.
 */
function scan(v) {
  v = (v || "").trim();
  if (!v || card) return;
  if (!S.master.length) {
    say("Impor master_barang.xlsx dulu.", true);
    return;
  }

  // Debounce pencegah duplikasi cepat barcode yang sama dalam tempo 2.5 detik
  const now = Date.now();
  if (v === last && now - lastT < 2500) return;
  last = v;
  lastT = now;

  const id = S.map[v];
  if (id && byId[id]) {
    // Jika opsi fast mode aktif, langsung catat qty +1 tanpa menampilkan popup konfirmasi
    if ($("fast").checked) {
      S.log.push({ t: now, bc: v, id, qty: 1 });
      say("+1 " + byId[id].nama);
      flash();
      after();
      return;
    }
    card = { bc: v, id };
    flash();
    renderCard();
  } else {
    // Barcode belum dikenal, buka kartu pencarian/pemetaan
    card = { bc: v, pending: true };
    flash();
    renderCard();
  }
}

/**
 * ==========================================================================
 * FUNGSI-FUNGSI PENGOLAHAN FILE EXCEL (IMPORT & EXPORT)
 * ==========================================================================
 */
/**
 * Membaca dan mengimpor file Excel master barang (.xlsx / .xls).
 * Mengurai kolom ID, kategori, supplier, satuan, kode barang, nama barang, serta relasi barcode jika tersedia.
 * Semua barcode yang ada di kolom "barcode" langsung diterima dan dipetakan ke ID barang.
 * @param {File} f - Objek file Excel dari elemen input type="file".
 */
async function importMaster(f) {
  try {
    const wb = XLSX.read(await f.arrayBuffer()),
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        defval: "",
      });

    if (!rows.length || !("nama_barang" in rows[0]) || !("id" in rows[0])) {
      say(
        "Kolom id dan nama_barang tidak ditemukan di sheet pertama.",
        true
      );
      return;
    }

    S.master = rows
      .filter((r) => r.nama_barang !== "" && r.id !== "")
      .map((r) => ({
        id: Number(r.id),
        kat: r.id_kategori,
        sup: r.id_supplier,
        sat: r.id_satuan,
        kode: String(r.kode_barang),
        nama: String(r.nama_barang),
      }));

    // Muat semua barcode dari kolom "barcode" di Excel, langsung dipetakan ke ID barang
    if ("barcode" in rows[0]) {
      S.map = {};
      let bcCount = 0;

      rows.forEach((r) => {
        // Konversi ke string untuk menangani barcode angka dari Excel
        const raw = String(r.barcode ?? "").trim();
        if (!raw || raw === "undefined" || raw === "null") return;

        raw
          .split("|")
          .map((x) => x.trim())
          .filter(Boolean)
          .forEach((b) => {
            S.map[b] = Number(r.id);
            bcCount++;
          });
      });

      index();
      say(
        "Master dimuat: " +
          S.master.length +
          " barang, " +
          bcCount +
          " barcode terdaftar."
      );
      after();
      return;
    }

    index();
    say("Master dimuat: " + S.master.length + " barang.");
    after();
  } catch (e) {
    say("File tidak bisa dibaca. Pastikan formatnya .xlsx.", true);
  }
}

/**
 * Membuat objek struktur data workbook Excel baru menggunakan pustaka XLSX.
 * Menghasilkan 2 lembar kerja (sheets):
 * 1. "master_barang": Daftar seluruh barang master beserta total akumulasi stok masuk dan daftar barcodenya.
 * 2. "log_scan": Riwayat detail setiap kali scan dilakukan (waktu, barcode, nama barang, jumlah).
 * @returns {Object} Objek Workbook XLSX siap unduh atau simpan.
 */
function workbook() {
  const stok = {},
    bcs = {};
  S.log.forEach((l) => (stok[l.id] = (stok[l.id] || 0) + l.qty));
  Object.entries(S.map).forEach(([b, id]) =>
    (bcs[id] = bcs[id] || []).push(b)
  );

  const a = [
    [
      "id",
      "id_kategori",
      "id_supplier",
      "id_satuan",
      "kode_barang",
      "nama_barang",
      "barcode",
      "stok_masuk",
    ],
  ];
  S.master.forEach((m) =>
    a.push([
      m.id,
      m.kat,
      m.sup,
      m.sat,
      m.kode,
      m.nama,
      (bcs[m.id] || []).join(" | "),
      stok[m.id] || "",
    ])
  );

  const l = [["waktu", "barcode", "kode_barang", "nama_barang", "jumlah"]];
  S.log.forEach((x) => {
    const m = byId[x.id] || {};
    l.push([fmt(x.t), x.bc, m.kode || "", m.nama || "", x.qty]);
  });

  const wb = XLSX.utils.book_new(),
    w1 = XLSX.utils.aoa_to_sheet(a),
    w2 = XLSX.utils.aoa_to_sheet(l);

  w1["!cols"] = [
    { wch: 6 },
    { wch: 11 },
    { wch: 11 },
    { wch: 9 },
    { wch: 12 },
    { wch: 44 },
    { wch: 22 },
    { wch: 11 },
  ];
  w2["!cols"] = [
    { wch: 22 },
    { wch: 18 },
    { wch: 12 },
    { wch: 44 },
    { wch: 8 },
  ];

  XLSX.utils.book_append_sheet(wb, w1, "master_barang");
  XLSX.utils.book_append_sheet(wb, w2, "log_scan");
  return wb;
}

/**
 * Menyimpan data secara otomatis langsung ke file Excel fisik di komputer pengguna
 * melalui File System Access API (jika pengguna sudah menghubungkan file via tombol "Hubungkan ke file Excel").
 */
async function saveLinked() {
  if (!handle) return;
  if (saving) {
    dirty = true;
    return;
  }
  saving = true;
  try {
    const out = XLSX.write(workbook(), { bookType: "xlsx", type: "array" });
    const w = await handle.createWritable();
    await w.write(out);
    await w.close();
    $("linked").textContent =
      "Tersimpan otomatis ke " +
      handle.name +
      " (" +
      new Date().toLocaleTimeString("id-ID") +
      ")";
  } catch (e) {
    $("linked").textContent =
      "Gagal menyimpan ke file. Tutup file itu di Excel, lalu scan lagi.";
  }
  saving = false;
  if (dirty) {
    dirty = false;
    saveLinked();
  }
}

/**
 * ==========================================================================
 * FUNGSI-FUNGSI PENGELOLAAN KAMERA & SCANNER
 * ==========================================================================
 */

/**
 * Memulai akses kamera perangkat, mendeteksi daftar kamera yang tersedia,
 * dan menjalankan timer loop pemindaian barcode (menggunakan native BarcodeDetector API
 * jika didukung browser, atau fallback ke library ZXing).
 * @param {string} [id] - DeviceId kamera tertentu yang dipilih dari dropdown.
 */
async function startCam(id) {
  stopCam();
  try {
    stream = await navigator.mediaDevices.getUserMedia({
      video: Object.assign(
        { width: { ideal: 1280 }, height: { ideal: 720 } },
        id ? { deviceId: { exact: id } } : { facingMode: "environment" }
      ),
    });
  } catch (e) {
    say(
      "Kamera tidak bisa dibuka. Izinkan akses kamera dan buka lewat https atau localhost.",
      true
    );
    return;
  }

  const v = $("v");
  v.srcObject = stream;
  await v.play();
  $("cam").textContent = "🛑 Matikan kamera";

  // Ambil daftar perangkat kamera untuk mengisi opsi dropdown kamera
  const devs = (await navigator.mediaDevices.enumerateDevices()).filter(
    (d) => d.kind === "videoinput"
  );
  if (devs.length > 1) {
    const s = $("dev");
    s.style.display = "";
    s.innerHTML = "";
    devs.forEach((d, i) =>
      s.add(new Option(d.label || "Kamera " + (i + 1), d.deviceId))
    );
    const cur = stream.getVideoTracks()[0].getSettings().deviceId;
    if (cur) s.value = cur;
  }

  const cv = document.createElement("canvas"),
    cx = cv.getContext("2d", { willReadFrequently: true });
  const fmts = [
    "ean_13",
    "ean_8",
    "upc_a",
    "upc_e",
    "code_128",
    "code_39",
    "code_93",
    "itf",
    "codabar",
    "qr_code",
  ];
  let bd = null,
    zx = null,
    busy = false;

  // Coba inisialisasi BarcodeDetector bawaan browser jika tersedia
  if ("BarcodeDetector" in window) {
    try {
      bd = new BarcodeDetector({ formats: fmts });
    } catch (e) {
      bd = null;
    }
  }

  // Fallback ke library ZXing jika BarcodeDetector native tidak tersedia
  const initZx = () => {
    if (zx || !window.ZXing) return;
    const hh = new Map();
    hh.set(ZXing.DecodeHintType.TRY_HARDER, true);
    zx = new ZXing.MultiFormatReader();
    zx.setHints(hh);
  };
  if (!bd) initZx();

  // Loop pemindaian frame video setiap 150ms
  timer = setInterval(async () => {
    if (busy || v.videoWidth === 0) return;
    busy = true;
    try {
      if (bd) {
        try {
          const r = await bd.detect(v);
          if (r.length) scan(r[0].rawValue);
        } catch (e) {
          bd = null;
          initZx();
        }
      } else if (zx) {
        const s = Math.min(1, 1280 / v.videoWidth);
        cv.width = v.videoWidth * s;
        cv.height = v.videoHeight * s;
        cx.drawImage(v, 0, 0, cv.width, cv.height);
        try {
          scan(
            zx
              .decode(
                new ZXing.BinaryBitmap(
                  new ZXing.HybridBinarizer(
                    new ZXing.HTMLCanvasElementLuminanceSource(cv)
                  )
                )
              )
              .getText()
          );
        } catch (e) {}
      }
    } catch (e) {}
    busy = false;
  }, 150);

  say("Kamera aktif. Arahkan ke barcode.");
}

/**
 * Menghentikan stream video kamera, mematikan track kamera,
 * mengosongkan elemen video, serta mematikan timer pemindaian.
 */
function stopCam() {
  clearInterval(timer);
  if (stream) {
    stream.getTracks().forEach((t) => t.stop());
    stream = null;
  }
  $("v").srcObject = null;
  $("cam").textContent = "📷 Mulai kamera";
}

/**
 * ==========================================================================
 * EVENT LISTENERS (INTERAKSI PENGGUNA)
 * ==========================================================================
 */

// Tombol nyalakan/matikan kamera
$("cam").onclick = () => (stream ? stopCam() : startCam());

// Ganti sumber kamera saat opsi di dropdown diubah
$("dev").onchange = (e) => startCam(e.target.value);

// Input barcode via scanner USB atau pengetikan manual (saat menekan tombol Enter)
$("man").addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    scan(e.target.value);
    e.target.value = "";
  }
});

// Tombol impor file master Excel
$("imp").onclick = () => $("file").click();
$("file").onchange = (e) => {
  if (e.target.files[0]) importMaster(e.target.files[0]);
  e.target.value = "";
};

// Tombol unduh file Excel hasil scan
$("dl").onclick = () => {
  if (!S.master.length) {
    say("Belum ada data.", true);
    return;
  }
  XLSX.writeFile(
    workbook(),
    "master_barang_" + new Date().toISOString().slice(0, 10) + ".xlsx"
  );
};

// Tombol kosongkan catatan log scan
$("clr").onclick = () => {
  if (
    S.log.length &&
    confirm("Hapus semua catatan scan? Pasangan barcode tetap tersimpan.")
  ) {
    S.log = [];
    last = "";
    after();
  }
};

// Tombol hubungkan ke file Excel lokal langsung (File System Access API)
$("link").onclick = async () => {
  if (!window.showSaveFilePicker) {
    say(
      "Simpan otomatis butuh Chrome atau Edge di komputer. Gunakan tombol Unduh .xlsx.",
      true
    );
    return;
  }
  try {
    handle = await showSaveFilePicker({
      suggestedName: "master_barang_hasil.xlsx",
      types: [
        {
          description: "Excel",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
              [".xlsx"],
          },
        },
      ],
    });
    await saveLinked();
    say("File terhubung. Setiap perubahan langsung ditulis ke Excel.");
  } catch (e) {}
};

/**
 * ==========================================================================
 * INISIALISASI AWAL SAAT HALAMAN DIBUKA
 * ==========================================================================
 */
index();
render();
if (S.master.length) say("Data sebelumnya dipulihkan dari browser.");
