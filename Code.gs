// =========================================================================
// BACKEND APPS SCRIPT INPUT FROM HTML
// =========================================================================

const FOLDER_BUKTI_TRANSFER = "1QyLKT6Iwed9BWC9GRDBaso-pWb4vADBX";
const FOLDER_BUKTI_IG_1 = "1O-N-pJMClwlxzKEQZCya80m3AhO3DNYw";
const FOLDER_BUKTI_IG_2 = "1q5pRNdsTCnBiOE1-GcJXMUKDAEwRNol-";
const FOLDER_BUKTI_TAG = "1a3pT2BcJnnGejRltQPV97pcRa5p46xY5";

const SHEET_DATA = "Data Pendaftar";
const SHEET_SETUP = "Setup Sistem";

function doOptions(e) {
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({ "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, GET, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" });
}

function doGet(e) { return prosesDataKirim(); }

function prosesDataKirim() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSetup = ss.getSheetByName(SHEET_SETUP);
  const sheetData = ss.getSheetByName(SHEET_DATA);

  const dataPengaturan = sheetSetup.getRange("B2:C" + (sheetSetup.getLastRow() || 2)).getValues();
  let settings = {};
  dataPengaturan.forEach(row => { if(row[0] && !row[0].includes("PIN")) settings[row[0]] = row[1]; });

  const snkText = settings["Isi_S&K"] || "S&K Belum diatur.";
  const lastRowData = sheetData.getLastRow();
  const totalPendaftar = lastRowData > 1 ? lastRowData - 1 : 0;

  return ContentService.createTextOutput(JSON.stringify({ status: "success", snk: snkText, pengaturan: settings, total_pendaftar: totalPendaftar })).setMimeType(ContentService.MimeType.JSON);
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
      // 1. ADMIN LOGIN & TARIK DATA PESERTA DEWASA (23 KOLOM)
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
          const dataRaw = sheetDataReg.getRange(2, 1, lastRow - 1, 23).getValues();
          dataRaw.forEach(row => {
            if (row[0] && row[1]) {
              let kuisGabungan = row[10] + "\n\n" + row[11] + "\n\n" + row[12] + "\n\n" + row[13];

              pendaftar.push({
                tanggal: row[0], orderId: row[1], email: row[2], hp: row[3],
                namaPeserta: row[4], gender: row[5], domisili: row[6],
                infoDari: row[7], temaPilihan: row[8], pendidikan: row[9],
                dataKuis: kuisGabungan, harga: row[14], statusBayar: row[15],
                bukti: row[16], logHadir: row[17], latitude: row[18], longitude: row[19],
                buktiFollow1: row[20], buktiFollow2: row[21], buktiTag3: row[22]
              });
            }
          });
        }
        pendaftar.reverse();
        return sendJSON({ status: "success", role: roleLogin, pendaftar: pendaftar, pengaturan: settingsObj });
      }

      // ===================================================================
      // 2. SCANNER QR KEHADIRAN
      // ===================================================================
      else if (data.action === "scan_qr") {
        const values = sheetDataReg.getDataRange().getValues();
        const qrCode = data.qr_code;
        let tz = ss.getSpreadsheetTimeZone();

        for (let i = 1; i < values.length; i++) {
          if (values[i][1] === qrCode) {
            let detailPeserta = {
              namaPeserta: values[i][4], gender: values[i][5], domisili: values[i][6],
              statusBayar: values[i][15],
              dataKuis: values[i][10] + "\n\n" + values[i][11] + "\n\n" + values[i][12] + "\n\n" + values[i][13],
              buktiFollow1: values[i][20], buktiFollow2: values[i][21], buktiTag3: values[i][22]
            };

            let statusBayarCek = values[i][15];
            let logHadirLama = values[i][17] || "";

            if (statusBayarCek === "UNPAID") return sendJSON({ status: "warning", message: `⚠️ Status Pembayaran masih UNPAID.\nAtas nama: ${detailPeserta.namaPeserta}` });
            if (logHadirLama !== "") return sendJSON({ status: "warning", message: `⚠️ PESERTA SUDAH CHECK-IN SEBELUMNYA.\nAtas nama: ${detailPeserta.namaPeserta}` });

            const d = new Date();
            const waktuSaja = Utilities.formatDate(d, tz, "HH:mm") + " WIB";
            sheetDataReg.getRange(i+1, 18).setValue(waktuSaja);

            return sendJSON({ status: "success", waktu: waktuSaja, peserta: detailPeserta });
          }
        }
        return sendJSON({ status: "error", message: "❌ QR Tidak Valid." });
      }

      // ===================================================================
      // 3. ADMIN: UPDATE STATUS PEMBAYARAN MANUAL
      // ===================================================================
      else if (data.action === "update_bayar") {
         const values = sheetDataReg.getDataRange().getValues();
         for(let i = 1; i < values.length; i++) {
            if(values[i][1] === data.orderId) { sheetDataReg.getRange(i+1, 16).setValue(data.status_baru); break; }
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
           if (typeof valToSave === 'string' && valToSave.startsWith('0')) valToSave = "'" + valToSave;
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
      // 5. ADMIN: HAPUS PESERTA
      // ===================================================================
      else if (data.action === "hapus_peserta") {
        const orderIdTarget = data.orderId;
        const values = sheetDataReg.getDataRange().getValues();
        let rowToDelete = -1;
        for (let i = values.length - 1; i >= 1; i--) { if (values[i][1] === orderIdTarget) { rowToDelete = i + 1; break; } }
        if (rowToDelete !== -1) { sheetDataReg.deleteRow(rowToDelete); return sendJSON({ status: "success" }); }
        else return sendJSON({ status: "error" });
      }

      // ===================================================================
      // 6. ADMIN: EDIT DATA PESERTA
      // ===================================================================
      else if (data.action === "edit_peserta") {
        const orderIdTarget = data.orderId;
        const dataBaru = data.dataBaru;
        const values = sheetDataReg.getDataRange().getValues();
        let rowToEdit = -1;
        for (let i = values.length - 1; i >= 1; i--) { if (values[i][1] === orderIdTarget) { rowToEdit = i + 1; break; } }
        if (rowToEdit !== -1) {
          sheetDataReg.getRange(rowToEdit, 5).setValue(dataBaru.namaPeserta);
          sheetDataReg.getRange(rowToEdit, 6).setValue(dataBaru.gender);
          sheetDataReg.getRange(rowToEdit, 10).setValue(dataBaru.pendidikan);
          sheetDataReg.getRange(rowToEdit, 7).setValue(dataBaru.domisili);
          sheetDataReg.getRange(rowToEdit, 4).setValue("'" + dataBaru.hp);
          return sendJSON({ status: "success" });
        } else return sendJSON({ status: "error" });
      }

      // ===================================================================
      // 7. LOGIKA PENDAFTARAN BARU (FORM SUBMIT)
      // ===================================================================
      else if (!data.action) {

        // --- AUTO HEADER GENERATOR BILA BARIS 1 KOSONG ---
        if (sheetDataReg.getRange("A1").getValue() === "") {
          sheetDataReg.getRange("A1:W1").setValues([[
            "Waktu Daftar", "Order ID", "Email", "No WA", "Nama Lengkap", "Gender", "Domisili",
            "Tahu Info Dari", "Tema Pilihan", "Pendidikan Terakhir",
            "1. Masalah yang paling menyita pikiran?", "2. Uang yang sudah dikeluarkan?", "3. Tujuan mempelajari hal tersebut?", "4. Harapan mengikuti kajian ini?",
            "Harga Tiket / Infaq", "Status Bayar", "Link Bukti Bayar", "Waktu Check-In", "Latitude", "Longitude",
            "Bukti Follow @al.mawaaizh", "Bukti Follow @ioubahasaindonesia", "Bukti Tag 3 Teman"
          ]]);
          sheetDataReg.getRange("A1:W1").setFontWeight("bold");
        }

        let fileUrl = "Tidak ada file";
        if (data.buktiBayarBase64) fileUrl = uploadImageToDrive(data.buktiBayarBase64, "Bukti_" + data.peserta[0].namaPeserta, FOLDER_BUKTI_TRANSFER);

        let buktiFollow1Url = "Tidak ada file"; let buktiFollow2Url = "Tidak ada file"; let buktiFollow3Url = "Tidak ada file";
        if (data.buktiFollow1Base64) buktiFollow1Url = uploadImageToDrive(data.buktiFollow1Base64, "Follow1_" + data.peserta[0].namaPeserta, FOLDER_BUKTI_IG_1);
        if (data.buktiFollow2Base64) buktiFollow2Url = uploadImageToDrive(data.buktiFollow2Base64, "Follow2_" + data.peserta[0].namaPeserta, FOLDER_BUKTI_IG_2);
        if (data.buktiFollow3Base64) buktiFollow3Url = uploadImageToDrive(data.buktiFollow3Base64, "Tag3_" + data.peserta[0].namaPeserta, FOLDER_BUKTI_TAG);

        const tglMasuk = new Date();
        const prefix = data.kodePrefix || "EVT";
        const thn = tglMasuk.getFullYear(); const bln = ("0" + (tglMasuk.getMonth() + 1)).slice(-2);
        const tgl = ("0" + tglMasuk.getDate()).slice(-2); const jam = ("0" + tglMasuk.getHours()).slice(-2);
        const mnt = ("0" + tglMasuk.getMinutes()).slice(-2); const dtk = ("0" + tglMasuk.getSeconds()).slice(-2);
        const baseTransactionId = prefix + "-" + thn + bln + tgl + "-" + jam + mnt + dtk + "-" + Math.floor(1000 + Math.random() * 9000);
        let generatedOrderIds = [];
        let hargaTiketFix = parseInt(data.infaqNominal) || 0;

        let q1 = "-"; let q2 = "-"; let q3 = "-"; let q4 = "-";
        try {
          if (data.dataKuisJson) {
            let kuisObj = JSON.parse(data.dataKuisJson);
            q1 = kuisObj.pertanyaan_1_masalah || "-";
            q2 = kuisObj.pertanyaan_2_pengeluaran || "-";
            q3 = kuisObj.pertanyaan_3_tujuan || "-";
            q4 = kuisObj.pertanyaan_4_harapan || "-";
          }
        } catch (e) {}

        data.peserta.forEach((p, index) => {
          let statusBayar = "UNPAID";
          let uniqueOrderId = baseTransactionId + "-" + (index + 1);
          generatedOrderIds.push(uniqueOrderId);

          let linkBuktiFix = index === 0 ? fileUrl : "Ikut Pendaftar Pertama";
          let linkFollow1Fix = index === 0 ? buktiFollow1Url : "Ikut Pendaftar Pertama";
          let linkFollow2Fix = index === 0 ? buktiFollow2Url : "Ikut Pendaftar Pertama";
          let linkFollow3Fix = index === 0 ? buktiFollow3Url : "Ikut Pendaftar Pertama";

          // EXACT 23 COLUMNS: A to W
          sheetDataReg.appendRow([
            tglMasuk,                           // A: Waktu Daftar
            uniqueOrderId,                      // B: Order ID Unik
            p.email,                            // C: Email
            "'" + p.hp,                         // D: No WA
            p.namaPeserta,                      // E: Nama Lengkap Peserta
            p.gender,                           // F: Gender
            p.domisili,                         // G: Domisili
            data.infoDari,                      // H: Info Dari
            data.temaPilihan,                   // I: Tema Pilihan
            data.pendidikanTerakhir,            // J: Pendidikan Terakhir
            q1,                                 // K: Masalah (Q1)
            q2,                                 // L: Pengeluaran (Q2)
            q3,                                 // M: Tujuan (Q3)
            q4,                                 // N: Harapan (Q4)
            hargaTiketFix,                      // O: HARGA TIKET / INFAQ
            statusBayar,                        // P: STATUS BAYAR
            linkBuktiFix,                       // Q: LINK BUKTI BAYAR
            "",                                 // R: WAKTU CHECK-IN
            p.latitude || "",                   // S: LATITUDE
            p.longitude || "",                  // T: LONGITUDE
            linkFollow1Fix,                     // U: BUKTI FOLLOW IG 1
            linkFollow2Fix,                     // V: BUKTI FOLLOW IG 2
            linkFollow3Fix                      // W: BUKTI TAG TEMAN
          ]);
        });

        return sendJSON({ status: "success", orderIds: generatedOrderIds, linkBukti: fileUrl });
      }

    } catch(err) { return sendJSON({ status: "error", message: err.toString() }); } finally { lock.releaseLock(); }
  } catch (fatalErr) { return sendJSON({ status: "error", message: fatalErr.toString() }); }
}

function sendJSON(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function uploadImageToDrive(base64Str, fileName, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const mimeType = base64Str.split(';')[0].split(':')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Str.split(',')[1]), mimeType, fileName + ".jpg");
  return folder.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW).getUrl();
}
