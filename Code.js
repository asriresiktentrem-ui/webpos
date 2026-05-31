const SPREADSHEET_ID = "1HQGcXniPIaNNp-VwYzO18_azqrC4z7HRHecJTXNxtng";
const FOLDER_ID = "13QVy1ux3PnLd6-WcA5_MeykGxqp6sjOL";

function doGet(e) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        // API BACKGROUND HANDLER (Dipanggil oleh verifikasi.html di GitHub)
        if (e.parameter.action === "api_verify") {
            const row = parseInt(e.parameter.row);
            const regSheet = ss.getSheetByName("Registrasi");
            
            if (!regSheet) {
                return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet Registrasi tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
            }
            
            // 1. Update status menjadi Terverifikasi
            regSheet.getRange(row, 2).setValue("Terverifikasi");
            
            // 2. Ambil data penyewa
            const namaPenyewa = regSheet.getRange(row, 3).getValue();
            const noKamar = regSheet.getRange(row, 4).getValue();
            let rawNoWa = regSheet.getRange(row, 6).getValue().toString().trim();
            
            // Bersihkan segala karakter non-digit kecuali tanda +
            let noWa = rawNoWa.replace(/[^0-9+]/g, "");
            if (noWa.startsWith("0")) {
                noWa = "62" + noWa.slice(1);
            } else if (!noWa.startsWith("62") && !noWa.startsWith("+62")) {
                noWa = "62" + noWa;
            }
            
            // 3. Ambil Password & SSID Wifi secara dinamis dari DataMaster (Kolom B & C)
            const masterSheet = ss.getSheetByName("DataMaster");
            let passwordWifi = "Belum Diatur";
            let ssidWifi = "Belum Diatur";
            
            if (masterSheet) {
                const lastRowMaster = masterSheet.getLastRow();
                if (lastRowMaster > 1) {
                    const masterData = masterSheet.getRange(2, 1, lastRowMaster - 1, 3).getValues();
                    for (let i = 0; i < masterData.length; i++) {
                        if (masterData[i][0].toString().trim() === noKamar.toString().trim()) {
                            passwordWifi = masterData[i][1].toString().trim();
                            if (masterData[i][2]) {
                                ssidWifi = masterData[i][2].toString().trim();
                            }
                            break;
                        }
                    }
                }
            }
            
            // 4. Teks Pesan WhatsApp Gaya Welcoming & Pemberitahuan Resmi
            const pesan = `Selamat Datang di Griya Ananda! ✨\n\nHalo *${namaPenyewa}*,\n\nKami dengan senang hati menginformasikan bahwa berkas registrasi hunian Anda untuk *Kamar ${noKamar}* telah resmi *TERVERIFIKASI* oleh sistem manajemen kos.\n\nSelamat bergabung menjadi bagian dari keluarga besar Griya Ananda. Semoga kenyamanan hunian ini mendukung segala produktivitas dan kelancaran aktivitas Anda ke depan.\n\nBerikut adalah detail akses fasilitas Wi-Fi area hunian Anda:\n- *SSID/Nama Wifi:* ${ssidWifi}\n- *Password Wifi:* ${passwordWifi}\n\nJika memerlukan bantuan lebih lanjut mengenai fasilitas hunian, silakan hubungi nomor Admin ini.\n\nTerima kasih dan selamat beristirahat! 🙏`;
            const waLink = "https://api.whatsapp.com/send?phone=" + noWa + "&text=" + encodeURIComponent(pesan);
            
            return ContentService.createTextOutput(JSON.stringify({ status: "success", waLink: waLink }))
                .setMimeType(ContentService.MimeType.JSON);
        }
        
        // JALUR AMBIL DAFTAR KAMAR (UNTUK DROPDOWN)
        const sheet = ss.getSheetByName("DataMaster");
        if (!sheet) {
            return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet 'DataMaster' tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
        }
        const lastRow = sheet.getLastRow();
        let daftarKamar = [];
        if (lastRow > 1) {
            const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            daftarKamar = values.map(row => row[0].toString().trim()).filter(kamar => kamar !== "");
        }
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: daftarKamar })).setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message })).setMimeType(ContentService.MimeType.JSON);
    }
}

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        if (data.pj_nama) {
            let sheet = ss.getSheetByName("Registrasi");
            if (!sheet) {
                sheet = ss.insertSheet("Registrasi");
                sheet.appendRow([
                    "Timestamp", "Status", "Nama Penyewa", "No Kamar", "Alamat Penyewa", "No WA Penyewa", 
                    "Nama Penanggung Jawab", "Alamat Penanggung Jawab", 
                    "Hubungan/Status", "No WA Penanggung Jawab", "Link Foto Identitas", "Aksi Verifikasi"
                ]);
            }
            
            let fileUrl = "";
            if (data.file_base64 && data.file_name) {
                const folder = DriveApp.getFolderById(FOLDER_ID);
                const decodedFile = Utilities.base64Decode(data.file_base64);
                const blob = Utilities.newBlob(decodedFile, data.file_type, data.file_name);
                const file = folder.createFile(blob);
                fileUrl = file.getUrl();
            }
            
            const webJembatanUrl = "https://asriresiktentrem-ui.github.io/webpos/verifikasi.html";
            const formulaVerifikasi = `=HYPERLINK("${webJembatanUrl}?row=" & ROW(); "🟢 VERIFIKASI & KIRIM WA")`;
            
            sheet.appendRow([
                new Date(), "Pending", data.nama, data.no_kamar, data.alamat, "'" + data.no_wa,
                data.pj_nama, data.pj_alamat, data.pj_status, "'" + data.pj_wa, fileUrl, formulaVerifikasi
            ]);
            
            return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Registrasi berhasil disimpan" })).setMimeType(ContentService.MimeType.JSON);
                
        } else {
            var sheet = ss.getSheetByName("Pembayaran");
            var folder = DriveApp.getFolderById(FOLDER_ID);
            var fileUrl = "";
            if (data.file_base64 && data.file_name) {
                var decodedFile = Utilities.base64Decode(data.file_base64);
                var blob = Utilities.newBlob(decodedFile, data.file_type, data.file_name);
                var file = folder.createFile(blob);
                fileUrl = file.getUrl();
            }
            sheet.appendRow([
                new Date(), data.nomor_kamar, data.nama_penyewa, data.jumlah_bayar, data.tanggal, fileUrl
            ]);
            return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data & Berkas Berhasil Disimpan" })).setMimeType(ContentService.MimeType.JSON);
        }
            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message })).setMimeType(ContentService.MimeType.JSON);
    }
}
