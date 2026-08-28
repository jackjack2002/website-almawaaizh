// =========================================================================
// BACKEND APPS SCRIPT - BRAVE KIDS PROJECT (V2 - HISTORICAL PRICE & MAP API)
// =========================================================================

const FOLDER_BUKTI = "1nqB9gJ2JpvzLmplnYRxLw00cwSkp4yVE"; 
const SHEET_DATA = "Data Pendaftar"; 
const SHEET_SETUP = "Setup Sistem";

function doOptions(e) {
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({ 
      "Access-Control-Allow-Origin": "*", 
      "Access-Control-Allow-Methods": "POST, GET, OPTIONS", 
      "Access-Control-Allow-Headers": "Content-Type" 
    });
}

function doGet(e) {
  return prosesDataKirim();
}

function prosesDataKirim() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSetup = ss.getSheetByName(SHEET_SETUP);
  const sheetData = ss.getSheetByName(SHEET_DATA); // Tambahan untuk menghitung kuota
   
  const dataPengaturan = sheetSetup.getRange("B2:C" + (sheetSetup.getLastRow() || 2)).getValues();
  let settings = {};
  dataPengaturan.forEach(row => {
    if(row[0] && !row[0].includes("PIN")) settings[row[0]] = row[1];
  });

  const snkText = settings["Isi_S&K"] || "S&K Belum diatur.";
  
  // Hitung berapa anak yang sudah terdaftar
  const lastRowData = sheetData.getLastRow();
  const totalPendaftar = lastRowData > 1 ? lastRowData - 1 : 0;

  const responseData = { status: "success", snk: snkText, pengaturan: settings, total_pendaftar: totalPendaftar };
   
  return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
     
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); 
     
    try {
      const sheetSetup = ss.getSheetByName(SHEET_SETUP);
      const sheetDataReg = ss.getSheetByName(SHEET_DATA);

      // ===================================================================
      // 1. ADMIN LOGIN & TARIK DATA (DASHBOARD & TABEL - 21 KOLOM MAPPER)
      // ===================================================================
      if (data.action === "login_admin" || data.action === "get_admin_data") {
        let pinSuperadmin = "123456"; let pinAdmin = "000000"; let settingsObj = {}; 
        const settingsData = sheetSetup.getRange("B2:C" + (sheetSetup.getLastRow() || 2)).getValues();
        settingsData.forEach(row => { 
          if(row[0] === "PIN_Superadmin") pinSuperadmin = row[1].toString(); 
          if(row[0] === "PIN_Admin") pinAdmin = row[1].toString(); 
          if(row[0]) settingsObj[row[0]] = row[1]; 
        });

        if (!settingsObj["PIN_Superadmin"]) settingsObj["PIN_Superadmin"] = pinSuperadmin;
        if (!settingsObj["PIN_Admin"]) settingsObj["PIN_Admin"] = pinAdmin;

        let roleLogin = "super_admin";
        if (data.action === "login_admin") {
          if (data.pin === pinSuperadmin) roleLogin = "super_admin";
          else if (data.pin === pinAdmin) roleLogin = "admin";
          else return sendJSON({ status: "error", message: "PIN Salah!" });
        }

        const lastRow = sheetDataReg.getLastRow(); 
        let pendaftar = [];
        if (lastRow > 1) { 
          // Membaca Kolom A sampai U (Total 21 Kolom)
          const dataRaw = sheetDataReg.getRange(2, 1, lastRow - 1, 21).getValues();
          dataRaw.forEach(row => {
            if (row[0] && row[1]) {
              pendaftar.push({
                tanggal: row[0], orderId: row[1], email: row[2], hp: row[3],
                namaAnak: row[4], panggilan: row[5], usia: row[6], sekolah: row[7], kelasSekolah: row[8],
                namaAba: row[9], namaUmma: row[10], domisili: row[11], penyakit: row[12], obat: row[13],
                infoDari: row[14], 
                harga: row[15],         // Kolom P (Harga Tiket Historis)
                statusBayar: row[16],   // Kolom Q (Status Bayar Bergeser)
                bukti: row[17],         // Kolom R (Link Bukti Bergeser)
                logHadir: row[18],      // Kolom S (Status Kehadiran Bergeser)
                latitude: row[19],      // Kolom T (Data Koordinat Peta)
                longitude: row[20]      // Kolom U (Data Koordinat Peta)
              });
            }
          });
        }
        pendaftar.reverse(); 
        return sendJSON({ status: "success", role: roleLogin, pendaftar: pendaftar, pengaturan: settingsObj });
      }

      // ===================================================================
      // 2. SCANNER QR (CHECK-IN KEHADIRAN - SINKRONISASI INDEKS BARU)
      // ===================================================================
      else if (data.action === "scan_qr") {
        const values = sheetDataReg.getDataRange().getValues();
        const qrCode = data.qr_code; // Berisi Order ID
        let tz = ss.getSpreadsheetTimeZone();

        for (let i = 1; i < values.length; i++) {
          if (values[i][1] === qrCode) { // Cek Kolom B (Order ID)
            let detailPeserta = {
              namaAnak: values[i][4], panggilan: values[i][5], usia: values[i][6],
              sekolah: values[i][7], penyakit: values[i][12], obat: values[i][13],
              statusBayar: values[i][16] // Membaca Kolom Q (Index 16)
            };

            let statusBayarCek = values[i][16]; // Kolom Q
            let logHadirLama = values[i][18] || ""; // Membaca Kolom S (Index 18)

            // CEK STATUS BAYAR: Tolak jika UNPAID
            if (statusBayarCek === "UNPAID") {
               return sendJSON({ status: "warning", message: `⚠️ Status Pembayaran masih UNPAID. Minta peserta konfirmasi ke Kasir!\nAtas nama: ${detailPeserta.namaAnak}` });
            }

            // CEK SUDAH HADIR ATAU BELUM
            if (logHadirLama !== "") {
               return sendJSON({ status: "warning", message: `⚠️ ANAK SUDAH CHECK-IN SEBELUMNYA.\nAtas nama: ${detailPeserta.namaAnak}` });
            }

            // JIKA LOLOS SEMUA: Catat Jam Hadir ke Kolom S (Ke-19)
            const d = new Date();
            const waktuSaja = Utilities.formatDate(d, tz, "HH:mm") + " WIB";
            sheetDataReg.getRange(i+1, 19).setValue(waktuSaja); 

            return sendJSON({ status: "success", waktu: waktuSaja, peserta: detailPeserta });
          }
        }
        return sendJSON({ status: "error", message: "❌ QR Tidak Valid. Data tidak ditemukan di database!" });
      }

      // ===================================================================
      // 3. ADMIN: UPDATE STATUS PEMBAYARAN MANUAL (KOLOM Q)
      // ===================================================================
      else if (data.action === "update_bayar") {
         const values = sheetDataReg.getDataRange().getValues();
         for(let i = 1; i < values.length; i++) {
            if(values[i][1] === data.orderId) {
               sheetDataReg.getRange(i+1, 17).setValue(data.status_baru); // Menargetkan Kolom Q (Index 17)
               break;
            }
         }
         return sendJSON({ status: "success" });
      }

      // ===================================================================
      // 4. ADMIN: SIMPAN PENGATURAN SISTEM
      // ===================================================================
      else if (data.action === "simpan_pengaturan") { 
         const keysData = sheetSetup.getRange("B2:B" + (sheetSetup.getLastRow() || 2)).getValues();
         const keysToUpdate = data.pengaturan; 
         for (let key in keysToUpdate) {
           let valToSave = keysToUpdate[key];
           if (typeof valToSave === 'string' && valToSave.startsWith('0')) { valToSave = "'" + valToSave; }
           let found = false;
           for (let i = 0; i < keysData.length; i++) {
             if (keysData[i][0] === key) { sheetSetup.getRange("C" + (i + 2)).setValue(valToSave); found = true; break; }
           }
           if (!found) { 
             let emptyRow = 2; while (sheetSetup.getRange("B" + emptyRow).getValue() !== "") emptyRow++;
             sheetSetup.getRange("A" + emptyRow).setValue("SISTEM"); sheetSetup.getRange("B" + emptyRow).setValue(key); sheetSetup.getRange("C" + emptyRow).setValue(valToSave);
           }
         }
         return sendJSON({ status: "success" });
      }

      // ===================================================================
      // 5. ADMIN: HAPUS PESERTA (TRASH)
      // ===================================================================
      else if (data.action === "hapus_peserta") {
        const orderIdTarget = data.orderId;
        const values = sheetDataReg.getDataRange().getValues();
        let rowToDelete = -1;
        
        // Cari baris berdasarkan Order ID (Kolom B / Index 1)
        for (let i = values.length - 1; i >= 1; i--) {
          if (values[i][1] === orderIdTarget) { 
            rowToDelete = i + 1; // Konversi index array ke nomor baris Google Sheets
            break;
          }
        }
        
        if (rowToDelete !== -1) {
          sheetDataReg.deleteRow(rowToDelete);
          return sendJSON({ status: "success", message: "Peserta dihapus." });
        } else {
          return sendJSON({ status: "error", message: "Peserta tidak ditemukan." });
        }
      }

      // ===================================================================
      // 6. ADMIN: EDIT PESERTA (FULL EDIT)
      // ===================================================================
      else if (data.action === "edit_peserta") {
        const orderIdTarget = data.orderId;
        const dataBaru = data.dataBaru;
        const values = sheetDataReg.getDataRange().getValues();
        let rowToEdit = -1;
        
        // Cari baris berdasarkan Order ID (Kolom B / Index 1)
        for (let i = values.length - 1; i >= 1; i--) {
          if (values[i][1] === orderIdTarget) { 
            rowToEdit = i + 1;
            break;
          }
        }
        
        if (rowToEdit !== -1) {
          // PROSES MENIMPA 11 DATA SEKALIGUS (Sesuaikan dengan urutan kolom Spreadsheet)
          sheetDataReg.getRange(rowToEdit, 5).setValue(dataBaru.namaAnak);       // Kolom E: Nama Anak
          sheetDataReg.getRange(rowToEdit, 6).setValue(dataBaru.panggilan);      // Kolom F: Panggilan
          sheetDataReg.getRange(rowToEdit, 7).setValue(dataBaru.usia);           // Kolom G: Usia (Tahun Lahir)
          sheetDataReg.getRange(rowToEdit, 8).setValue(dataBaru.sekolah);        // Kolom H: Sekolah
          sheetDataReg.getRange(rowToEdit, 9).setValue(dataBaru.kelasSekolah);   // Kolom I: Kelas Sekolah
          sheetDataReg.getRange(rowToEdit, 10).setValue(dataBaru.namaAba);       // Kolom J: Nama Aba
          sheetDataReg.getRange(rowToEdit, 11).setValue(dataBaru.namaUmma);      // Kolom K: Nama Umma
          sheetDataReg.getRange(rowToEdit, 12).setValue(dataBaru.domisili);      // Kolom L: Domisili
          sheetDataReg.getRange(rowToEdit, 4).setValue("'" + dataBaru.hp);       // Kolom D: No WA (Pakai petik)
          sheetDataReg.getRange(rowToEdit, 13).setValue(dataBaru.penyakit);      // Kolom M: Penyakit
          sheetDataReg.getRange(rowToEdit, 14).setValue(dataBaru.obat);          // Kolom N: Obat Pribadi
          
          return sendJSON({ status: "success", message: "Data berhasil diupdate secara menyeluruh." });
        } else {
          return sendJSON({ status: "error", message: "Peserta tidak ditemukan." });
        }
      }

      // ===================================================================
      // 7. LOGIKA PENDAFTARAN BARU (FORM SUBMIT - HARGA HISTORIS & KOORDINAT)
      // ===================================================================
      else if (!data.action) {
        let fileUrl = "Tidak ada file";
        if (data.buktiBayarBase64) fileUrl = uploadImageToDrive(data.buktiBayarBase64, "Bukti_" + data.peserta[0].namaAnak); 
        
        // Di file Code.gs, pada bagian: else if (!data.action) { ... }

const tglMasuk = new Date();
const baseTransactionId = "BKP-" + Math.floor(1000 + Math.random() * 9000);
let generatedOrderIds = [];

// Ambil nilai infaq dari frontend, tidak lagi membaca dari Setup Sistem
let hargaTiketFix = parseInt(data.infaqNominal) || 0; 

data.peserta.forEach((p, index) => {
  let statusBayar = "UNPAID"; 
  let uniqueOrderId = baseTransactionId + "-" + (index + 1);
  generatedOrderIds.push(uniqueOrderId);

  let linkBuktiFix = index === 0 ? fileUrl : "Ikut Anak Ke-1";

  // Kita gunakan kolom-kolom lama yang sudah tidak dipakai di frontend untuk menyimpan
  // data "Follow Munira" agar admin.html tetap aman berjalan.
  // Misalnya kita simpan Follow Munira di kolom "Kelas Sekolah" (Kolom I)

  sheetDataReg.appendRow([
    tglMasuk,              // A: Waktu Daftar
    uniqueOrderId,         // B: Order ID Unik
    p.email,               // C: Email
    "'" + p.hp,            // D: No WA
    p.namaAnak,            // E: Nama Lengkap
    "-",                   // F: Panggilan (dikosongkan)
    p.usia,                // G: Usia 
    "-",                   // H: Asal Sekolah (dikosongkan)
    "Follow Munira: " + p.followMunira, // I: Kelas Sekolah (Digunakan untuk data Follow)
    "-",                   // J: Nama Aba (dikosongkan)
    "-",                   // K: Nama Umma (dikosongkan)
    p.domisili,            // L: Domisili Teks
    "-",                   // M: Penyakit (dikosongkan)
    "-",                   // N: Obat (dikosongkan)
    data.infoDari,         // O: Info Dari
    hargaTiketFix,         // P: HARGA TIKET HISTORIS
    statusBayar,           // Q: STATUS BAYAR
    linkBuktiFix,          // R: LINK BUKTI
    "",                    // S: STATUS KEHADIRAN
    p.latitude || "",      // T: LATITUDE
    p.longitude || ""      // U: LONGITUDE
  ]);
});
return sendJSON({ status: "success", orderIds: generatedOrderIds, linkBukti: fileUrl });
      }
        
    } catch(err) { 
      return sendJSON({ status: "error", message: "Gagal memproses data: " + err.toString() }); 
    } finally { 
      lock.releaseLock(); 
    }
  } catch (fatalErr) {
    return sendJSON({ status: "error", message: "Fatal Server Error: " + fatalErr.toString() });
  }
}

function sendJSON(obj) { 
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); 
}

function uploadImageToDrive(base64Str, fileName) {
  const folder = DriveApp.getFolderById(FOLDER_BUKTI);
  const mimeType = base64Str.split(';')[0].split(':')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Str.split(',')[1]), mimeType, fileName + ".jpg");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
