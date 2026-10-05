const SHEET_ID = "1uwx9lszjrp_2SdgCQlCJyFg-KROF95crGuDFZQCzO-Y";
const SHEET_NAME = "Sheet1"; // Sesuaikan nama sheet Anda
const HEADERS    = [
    "Tanggal", "Nama", "Nominal", "Admin", "Denda",
    "Jumlah Bayar", "Status Bayar", "Keterangan",
    "No. WhatsApp",  
    "Terbayar"       
];

// ── Router GET ────────────────────────────────────────────────────────
function doGet(e) {
    try {
        const action = e.parameter.action || "getAll";
        if (action === "getAll")         return handleGetAll();
        if (action === "getNamesWithWA") return handleGetNamesWithWA();
        if (action === "getNames")       return handleGetNames();
        if (action === "getNotes")       return handleGetNotes();        // ← BARU
        return responseError("Action tidak dikenal: " + action);
    } catch (err) {
        return responseError("GET Error: " + err.message);
    }
}

function doPost(e) {
    try {
        const body   = JSON.parse(e.postData.contents);
        const action = body.action;
        if (action === "addTransaction") return handleAddTransaction(body);
        if (action === "updateStatus")   return handleUpdateStatus(body);
        if (action === "cicil")          return handleCicil(body);
        if (action === "addNote")        return handleAddNote(body);      // ← BARU
        if (action === "updateNote")     return handleUpdateNote(body);   // ← BARU
        if (action === "deleteNote")     return handleDeleteNote(body);   // ← BARU
        return responseError("Action tidak dikenal: " + action);
    } catch (err) {
        return responseError("POST Error: " + err.message);
    }
}

// ── GET: Semua transaksi ──────────────────────────────────────────────
// ── handleGetAll() — ganti 3 baris lama ──────────────────────────────
// HAPUS baris ini:
//   const tanggalTransaksi = new Date(tanggalRaw);
//   tanggalTransaksi.setHours(0, 0, 0, 0);
//   const today = new Date(); today.setHours(0, 0, 0, 0);
//
// GANTI dengan ini (taruh sekali di atas forEach):
function handleGetAll() {
    const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return responseSuccess({ transactions: [] });

    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    // ← Dua baris ini menggantikan semua new Date() + setHours() di bawah
    const today = hariIniWib();

    // Pass 1: cek siapa yang masih punya tunggakan
    const pelangganAdaBelumLunas = {};
    values.forEach(row => {
        const nama        = row[1];
        const statusBayar = row[6] || "BELUM LUNAS";
        if (!nama) return;
        if (pelangganAdaBelumLunas[nama] === undefined) {
            pelangganAdaBelumLunas[nama] = false;
        }
        if (statusBayar !== "LUNAS") pelangganAdaBelumLunas[nama] = true;
    });

    // Pass 2: hitung denda
    const transactions = [];
    values.forEach((row, index) => {
        const rowNumber   = index + 2;
        const tanggalRaw  = row[0];
        const nama        = row[1];
        const nominal     = Number(row[2]) || 0;
        const admin       = Number(row[3]) || 0;
        const statusBayar = row[6] || "BELUM LUNAS";
        const keterangan  = row[7] || "";
        const nomorWA     = row[8] || "";
        const terbayar    = Number(row[9]) || 0;

        if (!nama || !tanggalRaw) return;

        // ← Satu baris ini menggantikan new Date(tanggalRaw) + setHours()
        const tanggalTransaksi = toWibDate(tanggalRaw);

        const masihAdaYangNunggak = pelangganAdaBelumLunas[nama];

        let denda = 0;
        if (masihAdaYangNunggak) {
            denda = hitungDenda(tanggalTransaksi, today);
            const jumlahBayar = nominal + admin + denda;
            sheet.getRange(rowNumber, 5).setValue(denda);
            sheet.getRange(rowNumber, 6).setValue(jumlahBayar);
        } else {
            denda = Number(row[4]) || 0;
        }

        const jumlahBayar = nominal + admin + denda;
        const sisaTagihan = Math.max(0, jumlahBayar - terbayar);
        const tanggalFormatted = Utilities.formatDate(
            tanggalTransaksi, "Asia/Jakarta", "dd/MM/yyyy"
        );

        transactions.push({
            rowNumber, nama, nominal, admin, denda, jumlahBayar,
            tanggal:     tanggalFormatted,
            statusBayar, keterangan, nomorWA, terbayar, sisaTagihan
        });
    });

    return responseSuccess({ transactions });
}

// ── GET: Nama + Nomor WA (untuk AutoComplete + autofill) ─────────────
function handleGetNamesWithWA() {
    const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return responseSuccess({ customers: [] });

    // Baca kolom B (Nama) dan I (No. WA)
    const data = sheet.getRange(2, 2, lastRow - 1, 8).getValues();

    const customerMap = {};
    data.forEach(row => {
        const nama    = row[0];  // Kolom B
        const nomorWA = row[7] || ""; // Kolom I (index 7 dari B)
        // Simpan nomor WA terbaru jika nama sudah ada
        if (nama) customerMap[nama] = nomorWA;
    });

    const customers = Object.entries(customerMap)
        .map(([nama, nomorWA]) => ({ nama, nomorWA }))
        .sort((a, b) => a.nama.localeCompare(b.nama));

    return responseSuccess({ customers });
}

// ── GET: Nama saja (backup) ───────────────────────────────────────────
function handleGetNames() {
    const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return responseSuccess({ names: [] });

    const namaValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const namaunik   = [...new Set(namaValues.map(r => r[0]).filter(Boolean))].sort();
    return responseSuccess({ names: namaunik });
}

// ── POST: Tambah transaksi (+ nomorWA) ───────────────────────────────
function handleAddTransaction(body) {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!body.nama || !body.nominal || !body.tanggal)
        return responseError("Data tidak lengkap.");

    const nominal = Number(body.nominal);
    if (isNaN(nominal) || nominal <= 0) return responseError("Nominal tidak valid.");

    const admin   = hitungAdmin(nominal);
    const tanggal = toWibDate(body.tanggal);   // ← ganti parseTanggal
    const today   = hariIniWib();              // ← ganti new Date() + setHours
    const denda   = hitungDenda(tanggal, today);
    const jumlahBayar = nominal + admin + denda;
    const nomorWA     = body.nomorWA || "";
    const tanggalStr  = Utilities.formatDate(tanggal, "Asia/Jakarta", "dd/MM/yyyy");

    sheet.appendRow([
        tanggalStr, body.nama.trim(),
        nominal, admin, denda, jumlahBayar,
        "BELUM LUNAS", body.keterangan || "", nomorWA, 0
    ]);

    return responseSuccess({ message: "Transaksi berhasil disimpan." });
}

// ── POST: Update status (LUNAS/BELUM LUNAS) ───────────────────────────
function handleUpdateStatus(body) {
    if (!body.rowNumber) return responseError("rowNumber wajib.");
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    sheet.getRange(Number(body.rowNumber), 7).setValue(body.statusBayar || "LUNAS");
    return responseSuccess({ message: "Status diperbarui." });
}

// ── POST: Cicilan ─────────────────────────────────────────────────────
function handleCicil(body) {
    if (!body.rowNumber || !body.nominalCicil)
        return responseError("rowNumber dan nominalCicil wajib diisi.");

    const sheet        = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const rowNumber    = Number(body.rowNumber);
    const nominalCicil = Number(body.nominalCicil);

    if (nominalCicil <= 0) return responseError("Nominal cicilan harus lebih dari 0.");

    // Baca data baris ini
    const row          = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    const jumlahBayar  = Number(row[5]) || 0;
    const terbayarLama = Number(row[9]) || 0;
    const terbayarBaru = terbayarLama + nominalCicil;
    const sisa         = jumlahBayar - terbayarBaru;

    // Update kolom J (Terbayar)
    sheet.getRange(rowNumber, 10).setValue(terbayarBaru);

    // Append keterangan cicilan
    const tanggalHariIni  = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
    const keteranganLama  = row[7] || "";
    const catatanCicil    = "Cicil Rp" + nominalCicil.toLocaleString("id-ID") + " (" + tanggalHariIni + ")";
    const keteranganBaru  = keteranganLama ? keteranganLama + " | " + catatanCicil : catatanCicil;
    sheet.getRange(rowNumber, 8).setValue(keteranganBaru);

    // Auto LUNAS jika sisa sudah terlunasi
    const autoLunas = sisa <= 0;
    if (autoLunas) {
        sheet.getRange(rowNumber, 7).setValue("LUNAS");
    }

    return responseSuccess({
        message:      autoLunas ? "Cicilan masuk — tagihan LUNAS!" : "Cicilan berhasil dicatat.",
        terbayar:     terbayarBaru,
        sisa:         Math.max(0, sisa),
        autoLunas
    });
}

// ── Business logic & helpers (tidak berubah) ──────────────────────────
// ══════════════════════════════════════════════════════
// KONSTANTA — ganti di sini kalau aturan berubah lagi
// ══════════════════════════════════════════════════════
const GRACE_PERIOD_HARI = 1; // H+1: beli Senin → denda mulai Selasa

// ══════════════════════════════════════════════════════
// HELPER TANGGAL — timezone-safe, tidak pakai setHours
// ══════════════════════════════════════════════════════

// Konversi tanggal apapun (Date object ATAU string "dd/MM/yyyy")
// ke representasi UTC yang murni berdasarkan tanggal WIB-nya.
// Ini satu-satunya fungsi yang boleh membuat Date untuk kalkulasi denda.
function toWibDate(tanggalRaw) {
    let str;
    if (tanggalRaw instanceof Date) {
        // Date object dari sheet → paksa baca sebagai WIB dulu
        str = Utilities.formatDate(tanggalRaw, "Asia/Jakarta", "dd/MM/yyyy");
    } else {
        str = String(tanggalRaw); // Sudah string "dd/MM/yyyy"
    }

    // Parse "dd/MM/yyyy" → Date.UTC (tidak ada timezone ambiguity)
    const parts = str.split("/");
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1; // bulan 0-based
    const y = parseInt(parts[2], 10);
    return new Date(Date.UTC(y, m, d));
}

// Hari ini dalam WIB, direpresentasikan sebagai Date UTC
function hariIniWib() {
    const str = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");
    const parts = str.split("/");
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    return new Date(Date.UTC(y, m, d));
}

function hitungAdmin(nominal) {
    // Sesuaikan kembali dengan aturan admin bisnisnya April
    if (nominal >= 50000) {
        return 5000;
    } else if (nominal >= 20000) {
        return 3000;
    } else {
        return 2000; // Admin default jika nominal kecil
    }
}

// ══════════════════════════════════════════════════════
// HITUNG DENDA — bersih, tidak ada timezone issue
// ══════════════════════════════════════════════════════
function hitungDenda(tanggalTransaksi, hariIni) {
    // Keduanya WAJIB hasil dari toWibDate() atau hariIniWib()
    // sehingga comparison-nya selalu apples-to-apples
    const selisihMs     = hariIni.getTime() - tanggalTransaksi.getTime();
    const selisihHari   = Math.floor(selisihMs / 86400000);
    const hariKenaDenda = selisihHari - GRACE_PERIOD_HARI;

    if (hariKenaDenda < 0) return 0;
    return (hariKenaDenda + 1) * 1000;
}

// ══════════════════════════════════════════════════════
// parseTanggal — sekarang pakai toWibDate supaya konsisten
// ══════════════════════════════════════════════════════
function parseTanggal(str) {
    return toWibDate(str);
}


function parseTanggal(str) {
    if (str.includes("/")) {
        const [d, m, y] = str.split("/");
        return new Date(Number(y), Number(m) - 1, Number(d));
    }
    return new Date(str);
}

function responseSuccess(data) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", ...data }))
        .setMimeType(ContentService.MimeType.JSON);
}

function responseError(message) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message }))
        .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
// FITUR CATATAN — Sheet "Catatan" auto-dibuat jika belum ada
// ═══════════════════════════════════════════════════════════
const NOTES_SHEET_NAME = "Catatan";
const NOTES_HEADERS    = ["ID", "Judul", "Isi", "Tanggal Dibuat", "Tanggal Diubah"];

// ── Ambil sheet Catatan, buat otomatis kalau belum ada ────────────────
function getNotesSheet() {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(NOTES_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(NOTES_SHEET_NAME);
        sheet.getRange(1, 1, 1, NOTES_HEADERS.length).setValues([NOTES_HEADERS]);
        sheet.getRange(1, 1, 1, NOTES_HEADERS.length)
            .setFontWeight("bold")
            .setBackground("#1A1A1A")
            .setFontColor("#FFFFFF");
        sheet.setColumnWidth(1, 220); // Kolom ID (UUID) agak lebar
        sheet.setColumnWidth(3, 350); // Kolom Isi dilebarkan
    }
    return sheet;
}

// ── GET: Semua catatan ─────────────────────────────────────────────────
function handleGetNotes() {
    const sheet   = getNotesSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return responseSuccess({ notes: [] });

    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const notes  = [];

    values.forEach(row => {
        const id = row[0];
        if (!id) return; // Skip baris kosong

        notes.push({
            noteId:        id,
            judul:         row[1] || "",
            isi:           row[2] || "",
            tanggalDibuat: row[3] || "",
            tanggalDiubah: row[4] || ""
        });
    });

    return responseSuccess({ notes });
}

// ── POST: Tambah catatan baru ────────────────────────────────────────
function handleAddNote(body) {
    if (!body.noteId) {
        return responseError("noteId wajib diisi.");
    }
    if (!body.judul && !body.isi) {
        return responseError("Catatan tidak boleh kosong.");
    }

    const sheet = getNotesSheet();
    sheet.appendRow([
        body.noteId,
        body.judul || "",
        body.isi   || "",
        body.tanggalDibuat || Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm"),
        body.tanggalDiubah || Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm")
    ]);

    return responseSuccess({ message: "Catatan berhasil disimpan." });
}

// ── POST: Update catatan ─────────────────────────────────────────────
function handleUpdateNote(body) {
    if (!body.noteId) return responseError("noteId wajib diisi.");

    const sheet  = getNotesSheet();
    const rowIdx = cariBarisCatatan(sheet, body.noteId);
    if (rowIdx === -1) return responseError("Catatan tidak ditemukan.");

    if (body.judul !== undefined) sheet.getRange(rowIdx, 2).setValue(body.judul);
    if (body.isi   !== undefined) sheet.getRange(rowIdx, 3).setValue(body.isi);
    sheet.getRange(rowIdx, 5).setValue(
        body.tanggalDiubah || Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm")
    );

    return responseSuccess({ message: "Catatan berhasil diperbarui." });
}

// ── POST: Hapus catatan ───────────────────────────────────────────────
function handleDeleteNote(body) {
    if (!body.noteId) return responseError("noteId wajib diisi.");

    const sheet  = getNotesSheet();
    const rowIdx = cariBarisCatatan(sheet, body.noteId);
    if (rowIdx === -1) return responseError("Catatan tidak ditemukan.");

    sheet.deleteRow(rowIdx);
    return responseSuccess({ message: "Catatan berhasil dihapus." });
}

// ── Helper: Cari nomor baris berdasarkan noteId (UUID) ─────────────────
function cariBarisCatatan(sheet, noteId) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return -1;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === noteId) return i + 2; // +2 karena data mulai baris 2
    }
    return -1;
}
