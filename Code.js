const SPREADSHEET_ID = "1HQGcXniPIaNNp-VwYzO18_azqrC4z7HRHecJTXNxtng";
const FOLDER_ID = "13QVy1ux3PnLd6-WcA5_MeykGxqp6sjOL";

function doGet(e) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        
        // JALUR A: API Verifikasi Registrasi Penghuni Baru
        if (e.parameter.action === "api_verify") {
            const row = parseInt(e.parameter.row);
            const regSheet = ss.getSheetByName("Registrasi");
            
            if (!regSheet) {
                return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet Registrasi tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
            }
            
            regSheet.getRange(row, 2).setValue("Terverifikasi");
            
            const namaPenyewa = regSheet.getRange(row, 3).getValue();
            const noKamar = regSheet.getRange(row, 4).getValue();
            let rawNoWa = regSheet.getRange(row, 6).getValue().toString().trim();
            
            let noWa = rawNoWa.replace(/[^0-9+]/g, "");
            if (noWa.startsWith("0")) {
                noWa = "62" + noWa.slice(1);
            } else if (!noWa.startsWith("62") && !noWa.startsWith("+62")) {
                noWa = "62" + noWa;
            }
            
            const masterSheet = ss.getSheetByName("DataMaster");
            let passwordWifi = "Belum Diatur";
            let ssidWifi = "Belum Diatur";
            
            if (masterSheet) {
                const lastRowMaster = masterSheet.getLastRow();
                if (lastRowMaster > 1) {
                    const masterData = masterSheet.getRange(2, 1, lastRowMaster - 1, 4).getValues();
                    for (let i = 0; i < masterData.length; i++) {
                        if (masterData[i][0].toString().trim() === noKamar.toString().trim()) {
                            passwordWifi = masterData[i][2].toString().trim();
                            ssidWifi = masterData[i][3].toString().trim();
                            break;
                        }
                    }
                }
            }
            
            const pesan = `Selamat Datang di Griya Ananda! ✨\n\nHalo *${namaPenyewa}*,\n\nKami dengan senang hati menginformasikan bahwa berkas registrasi hunian Anda untuk *Kamar ${noKamar}* telah resmi *TERVERIFIKASI* oleh sistem manajemen kos.\n\nSelamat bergabung menjadi bagian dari keluarga besar Griya Ananda. Semoga kenyamanan hunian ini mendukung segala produktivitas dan kelancaran aktivitas Anda ke depan.\n\nBerikut adalah detail akses fasilitas Wi-Fi area hunian Anda:\n- *SSID/Nama Wifi:* ${ssidWifi}\n- *Password Wifi:* ${passwordWifi}\n\nJika memerlukan bantuan lebih lanjut mengenai fasilitas hunian, silakan hubungi nomor Admin ini.\n\nTerima kasih dan selamat beristirahat! 🙏`;
            const waLink = "https://api.whatsapp.com/send?phone=" + noWa + "&text=" + encodeURIComponent(pesan);
            
            return ContentService.createTextOutput(JSON.stringify({ status: "success", waLink: waLink })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // JALUR B: API Verifikasi Pembayaran Kos + Return Detail Meta Kuitansi
        if (e.parameter.action === "api_verify_payment") {
            const row = parseInt(e.parameter.row);
            const paySheet = ss.getSheetByName("Pembayaran");
            
            if (!paySheet) {
                return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet Pembayaran tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
            }
            
            paySheet.getRange(row, 2).setValue("Terverifikasi");
            
            const noKamar = paySheet.getRange(row, 3).getValue().toString().trim();
            const namaPenyewa = paySheet.getRange(row, 4).getValue();
            let rawNoWa = paySheet.getRange(row, 5).getValue().toString().trim();
            const periode = paySheet.getRange(row, 6).getValue();
            const jumlahBayar = paySheet.getRange(row, 7).getValue();
            const tanggalTx = paySheet.getRange(row, 8).getValue();
            
            let noWa = rawNoWa.replace(/[^0-9+]/g, "");
            if (noWa.startsWith("0")) {
                noWa = "62" + noWa.slice(1);
            } else if (!noWa.startsWith("62") && !noWa.startsWith("+62")) {
                noWa = "62" + noWa;
            }
            
            const totalFormat = "Rp " + parseInt(jumlahBayar).toLocaleString('id-ID');
            const invoiceId = "INV-" + Date.now().toString().slice(-6);
            const webJembatanBayarUrl = "https://asriresiktentrem-ui.github.io/webpos/verifikasi_bayar.html?row=" + row;
            
            const pesan = `KUITANSI DIGITAL RESMI - GRIYA ANANDA ✨\n\nTerima kasih *${namaPenyewa}*,\n\nPembayaran kontribusi sewa kos Anda telah kami terima dan dinyatakan *TERVERIFIKASI* oleh Pemilik Kos.\n\nBerikut adalah rincian tanda terima pembayaran Anda:\n- *ID Kuitansi:* ${invoiceId}\n- *Nomor Kamar:* Kamar ${noKamar}\n- *Periode Sewa:* ${periode} Bulan\n- *Total Pembayaran:* ${totalFormat}\n- *Tanggal Transfer:* ${tanggalTx}\n- *Status:* LUNAS / TERVERIFIKASI\n\nTanda bukti transaksi ini sah. Anda dapat melihat, menyimpan dalam bentuk cetak dokumen GAMBAR (.PNG) atau berkas (.PDF) resmi melalui tautan kuitansi online kami berikut:\n👉 ${webJembatanBayarUrl}\n\nTerima kasih atas kedisiplinan Anda. Salam hangat! 🙏`;
            const waLink = "https://api.whatsapp.com/send?phone=" + noWa + "&text=" + encodeURIComponent(pesan);
            
            return ContentService.createTextOutput(JSON.stringify({ 
                status: "success", 
                waLink: waLink,
                details: {
                    invoice_id: invoiceId,
                    kamar: noKamar,
                    penyewa: namaPenyewa,
                    periode: periode + " Bulan",
                    total: totalFormat,
                    tanggal: tanggalTx
                }
            })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // JALUR C: Sinkronisasi Daftar Kamar, Harga, Nama & No WA Penyewa Aktif (Fix Bug M1-M10)
        if (e.parameter.action === "getKamar") {
            const masterSheet = ss.getSheetByName("DataMaster");
            const regSheet = ss.getSheetByName("Registrasi");
            
            if (!masterSheet) {
                return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet DataMaster tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);
            }
            
            const lastRowMaster = masterSheet.getLastRow();
            let resultData = [];
            
            if (lastRowMaster > 1) {
                const masterValues = masterSheet.getRange(2, 1, lastRowMaster - 1, 2).getValues();
                
                let regValues = [];
                if (regSheet) {
                    const lastRowReg = regSheet.getLastRow();
                    if (lastRowReg > 1) {
                        regValues = regSheet.getRange(2, 1, lastRowReg - 1, 6).getValues();
                    }
                }
                
                for (let i = 0; i < masterValues.length; i++) {
                    const kamar = masterValues[i][0].toString().trim();
                    const harga = parseInt(masterValues[i][1]) || 0;
                    
                    if (kamar === "") continue;
                    
                    let namaPenyewa = "Belum Ada Penghuni";
                    let noWaPenyewa = "";
                    
                    for (let j = regValues.length - 1; j >= 0; j--) {
                        const regStatus = regValues[j][1].toString().trim();
                        const regKamar = regValues[j][3].toString().trim();
                        
                        // Perbaikan Krusial: Paksa komparasi berbasis String murni agar M1 dan angka normal cocok sempurna
                        if (regKamar === kamar && regStatus === "Terverifikasi") {
                            namaPenyewa = regValues[j][2].toString().trim();
                            noWaPenyewa = regValues[j][5].toString().trim();
                            break;
                        }
                    }
                    
                    resultData.push({
                        kamar: kamar,
                        harga: harga,
                        penyewa: namaPenyewa,
                        no_wa: noWaPenyewa
                    });
                }
            }
            return ContentService.createTextOutput(JSON.stringify({ status: "success", data: resultData })).setMimeType(ContentService.MimeType.JSON);
        }
        
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
                new Date(), "Pending", data.nama, data.no_kamar.toString().trim(), data.alamat, "'" + data.no_wa,
                data.pj_nama, data.pj_alamat, data.pj_status, "'" + data.pj_wa, fileUrl, formulaVerifikasi
            ]);
            
            return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Registrasi berhasil disimpan" })).setMimeType(ContentService.MimeType.JSON);
                
        } else {
            let sheet = ss.getSheetByName("Pembayaran");
            if (!sheet) {
                sheet = ss.insertSheet("Pembayaran");
                sheet.appendRow([
                    "Timestamp", "Status", "Nomor Kamar", "Nama Penyewa", "No WA Penyewa", 
                    "Periode", "Jumlah Bayar", "Tanggal Transfer", "Link Bukti Transfer", "Aksi Verifikasi"
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
            
            const webJembatanBayarUrl = "https://asriresiktentrem-ui.github.io/webpos/verifikasi_bayar.html";
            const formulaVerifikasi = `=HYPERLINK("${webJembatanBayarUrl}?row=" & ROW(); "🟢 VERIFIKASI & BUKTI")`;
            
            sheet.appendRow([
                new Date(),
                "Pending",
                data.nomor_kamar.toString().trim(),
                data.nama_penyewa,
                "'" + data.no_wa,
                data.periode,
                data.jumlah_bayar,
                data.tanggal,
                fileUrl,
                formulaVerifikasi
            ]);
            
            return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data & Berkas Berhasil Disimpan" })).setMimeType(ContentService.MimeType.JSON);
        }
            
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message })).setMimeType(ContentService.MimeType.JSON);
    }
}
