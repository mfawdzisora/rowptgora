const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const axios = require('axios');

const DATA_FILE = path.join(__dirname, 'data.json');
const FOTO_DIR = path.join(__dirname, 'foto');

if (!fs.existsSync(FOTO_DIR)) fs.mkdirSync(FOTO_DIR);

function loadData() {
    try {
        if (fs.existsSync(DATA_FILE)) {
            let raw = fs.readFileSync(DATA_FILE, 'utf8');
            return JSON.parse(raw);
        }
    } catch (e) {
        console.log("⚠️ Gagal load data:", e.message);
    }
    return { laporan: [], antrian: [], laporanHujan: [] };
}

function saveData() {
    try {
        fs.writeFileSync(DATA_FILE, JSON.stringify({ laporan, antrian, laporanHujan }, null, 2));
    } catch (e) {
        console.log("⚠️ Gagal save data:", e.message);
    }
}

let { laporan, antrian, laporanHujan } = loadData();
if (!laporanHujan) laporanHujan = [];
let userState = {};
let verifiedUsers = {};
let loginState = {};

const token = '8637299952:AAFHGRpVDHBtoeJoCBT5yvs0OKSXNgEHTDk';

const bot = new TelegramBot(token, {
    polling: { interval: 3000, autoStart: true, params: { timeout: 10 } }
});
// Pastikan tidak ada getUpdates yang overlap saat start
bot.deleteWebHook().then(() => {
    console.log("✅ Webhook cleared, polling dimulai");
}).catch(err => {
    console.log("⚠️ deleteWebHook error:", err.message);
});

bot.on('polling_error', (err) => console.log(`⚠️ Polling error: ${err.code} — ${err.message}`));
bot.on('error', (err) => console.log(`⚠️ Bot error: ${err.message}`));

const kodePetugas = "PTG2026";
const kodeManajemen = "ADMIN2026";

async function kirimPesan(chatId, text, options = {}, retry = 3) {
    for (let i = 0; i < retry; i++) {
        try {
            return await bot.sendMessage(chatId, text, options);
        } catch (err) {
            console.log(`⚠️ Gagal kirim pesan (percobaan ${i+1}): ${err.message}`);
            if (i < retry - 1) await delay(2000);
        }
    }
    antrian.push({ chatId, text, options, waktu: new Date().toISOString() });
    saveData();
}

async function kirimDokumen(chatId, filePath, retry = 3) {
    for (let i = 0; i < retry; i++) {
        try {
            return await bot.sendDocument(chatId, filePath);
        } catch (err) {
            console.log(`⚠️ Gagal kirim dokumen (percobaan ${i+1}): ${err.message}`);
            if (i < retry - 1) await delay(2000);
        }
    }
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

setInterval(async () => {
    if (antrian.length === 0) return;
    let gagal = [];
    for (let item of antrian) {
        try {
            await bot.sendMessage(item.chatId, item.text, item.options);
        } catch (err) { gagal.push(item); }
    }
    antrian = gagal;
    saveData();
}, 30000);

async function downloadFoto(fileId, retry = 3) {
    for (let i = 0; i < retry; i++) {
        try {
            const fileLink = await bot.getFileLink(fileId);
            const fileName = `foto_${Date.now()}.jpg`;
            const filePath = path.join(FOTO_DIR, fileName);
            const response = await axios({ url: fileLink, method: 'GET', responseType: 'stream', timeout: 15000 });
            const writer = fs.createWriteStream(filePath);
            response.data.pipe(writer);
            await new Promise((resolve, reject) => {
                writer.on('finish', resolve);
                writer.on('error', reject);
            });
            return filePath;
        } catch (err) {
            console.log(`⚠️ Gagal download foto (percobaan ${i+1}): ${err.message}`);
            if (i < retry - 1) await delay(3000);
        }
    }
    return null;
}

async function downloadVideo(fileId, retry = 3) {
    for (let i = 0; i < retry; i++) {
        try {
            const fileLink = await bot.getFileLink(fileId);
            const fileName = `video_${Date.now()}.mp4`;
            const filePath = path.join(FOTO_DIR, fileName);
            const response = await axios({ url: fileLink, method: 'GET', responseType: 'stream', timeout: 30000 });
            const writer = fs.createWriteStream(filePath);
            response.data.pipe(writer);
            await new Promise((resolve, reject) => {
                writer.on('finish', resolve);
                writer.on('error', reject);
            });
            return filePath;
        } catch (err) {
            console.log(`⚠️ Gagal download video (percobaan ${i+1}): ${err.message}`);
            if (i < retry - 1) await delay(3000);
        }
    }
    return null;
}

// ==================
// EMBED FOTO KE EXCEL — LANGSUNG TANPA RESIZE
// ==================
async function embedFoto(workbook, sheet, fotoPath, colIndex, rowIndex) {
    try {
        if (!fotoPath || !fs.existsSync(fotoPath)) {
            console.log(`⚠️ Skip embed foto — file tidak ada: ${fotoPath}`);
            return;
        }
        const imageId = workbook.addImage({ filename: fotoPath, extension: 'jpeg' });
        sheet.addImage(imageId, {
            tl: { col: colIndex, row: rowIndex },
            ext: { width: 100, height: 75 }
        });
    } catch (e) {
        console.log(`⚠️ Gagal embed foto: ${e.message}`);
    }
}

// ==================
// HELPER HITUNG METER
// ==================
function hitungMeterSesi(sesi) {
    return sesi.reduce((total, s) => total + (s.akhir - s.awal), 0);
}

// ==================
// RESET JAM 12 MALAM
// ==================
function jadwalReset() {
    const now = new Date();
    
    // Hitung jam 00:00 WIB berikutnya (WIB = UTC+7)
    // 00:00 WIB = 17:00 UTC hari sebelumnya
    const nowUTC = now.getTime();
    const wibOffset = 7 * 60 * 60 * 1000; // UTC+7 dalam ms
    
    const nowWIB = new Date(nowUTC + wibOffset);
    
    // Besok jam 00:00 WIB
    const besokWIB = new Date(nowWIB);
    besokWIB.setUTCDate(nowWIB.getUTCDate() + 1);
    besokWIB.setUTCHours(0, 0, 0, 0);
    
    // Konversi balik ke UTC untuk setTimeout
    const targetUTC = besokWIB.getTime() - wibOffset;
    const selisih = targetUTC - nowUTC;

    setTimeout(() => {
        const tanggalBackup = nowWIB.toISOString().slice(0, 10); // YYYY-MM-DD WIB
        console.log("🔄 RESET DATA TENGAH MALAM WIB...");
        const backupFile = path.join(__dirname, `backup_${tanggalBackup}.json`);
        fs.writeFileSync(backupFile, JSON.stringify({ laporan, antrian, laporanHujan }, null, 2));
        laporan = [];
        antrian = [];
        laporanHujan = [];
        userState = {};
        saveData();
        console.log("✅ Data berhasil direset (WIB)");
        jadwalReset();
    }, selisih);

    const menitLagi = Math.round(selisih / 1000 / 60);
    console.log(`⏰ Reset WIB dijadwalkan dalam ${menitLagi} menit`);
}

function getTanggal() {
    // Pakai waktu WIB (UTC+7)
    const wibOffset = 7 * 60 * 60 * 1000;
    const d = new Date(Date.now() + wibOffset);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
}

// ==================
// FUNGSI BUAT SESI DARI LIST KP
// ==================
function buatSesi(kpList) {
    let kpAngka = kpList.map(kp => {
        let [km, m] = kp.split("+");
        return parseInt(km) * 1000 + parseInt(m);
    });
    kpAngka.sort((a, b) => a - b);

    let sesi = [];
    let sesiAwal = kpAngka[0], sesiAkhir = kpAngka[0], sesiTitik = 1;
    for (let i = 1; i < kpAngka.length; i++) {
        if (kpAngka[i] - kpAngka[i-1] <= 100) {
            sesiAkhir = kpAngka[i]; sesiTitik++;
        } else {
            sesi.push({ awal: sesiAwal, akhir: sesiAkhir, titik: sesiTitik });
            sesiAwal = kpAngka[i]; sesiAkhir = kpAngka[i]; sesiTitik = 1;
        }
    }
    sesi.push({ awal: sesiAwal, akhir: sesiAkhir, titik: sesiTitik });
    return sesi;
}

// ==================
// MENU
// ==================
function showMenu(chatId) {
    let role = verifiedUsers[chatId]?.role;
    let menu = [["📷 Dokumentasi Penyisiran Jalur ROW"]];

    if (role === "petugas") {
        menu.push(["📋 History Penyisiran"]);
        menu.push(["🗺 Real-Time Penyisiran All Area"]);
        menu.push(["🔀 Pindah Segment"]);
        menu.push(["🌧 Report Hujan"]);
        menu.push(["📤 Export Penyisiran Jalur ROW"]);
        menu.push(["🚪 Log out"]);
    }

    if (role === "admin") {
        menu.push(["📤 Export Penyisiran Jalur ROW"]);
        menu.push(["📊 Dashboard Live Penyisiran Jalur ROW"]);
        menu.push(["👷 Data Petugas Jalur ROW"]);
        menu.push(["🌿 Data Jalur ROW Bersemak"]);
        menu.push(["🚧 Data Finding/Pelanggaran Jalur ROW"]);
        menu.push(["🚪 Log out"]);
    }

    kirimPesan(chatId, "📋 MENU UTAMA", {
        reply_markup: { keyboard: menu, resize_keyboard: true }
    });
}

bot.onText(/\/start/, (msg) => {
    const chatId = msg.chat.id;
    loginState[chatId] = { step: "pilih_role" };
    kirimPesan(chatId, "🔐 Pilih jenis login:", {
        reply_markup: {
            keyboard: [["👷Login Petugas"], ["🧑‍💼Login Manajemen"]],
            resize_keyboard: true
        }
    });
});

bot.on('location', async (msg) => {
    const chatId = msg.chat.id;
    if (!userState[chatId]) return;

    const lat = msg.location.latitude;
    const lon = msg.location.longitude;

    // ==================
    // LOKASI UNTUK LAPORAN HUJAN
    // ==================
    if (userState[chatId].mode === "tunggu_lokasi_hujan") {
        userState[chatId].koordinatHujan = { lat, lon };
        await kirimPesan(chatId, `📍 Lokasi diterima ✅\n🌐 ${lat}, ${lon}\n\n⏳ Menyimpan laporan hujan...`);
        return prosesHasilHujan(chatId);
    }

    userState[chatId].koordinat = { lat, lon };

    await kirimPesan(chatId, `📍 Lokasi diterima ✅\n🌐 ${lat}, ${lon}\n\n⏳ Menyimpan data...`);
    userState[chatId].mode = null;
    return prosesHasilFoto(chatId);
});

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;
    const text = msg.text;
    if (!text) return;

    if (text === "👷Login Petugas") {
        loginState[chatId] = { step: "login_petugas" };
        return kirimPesan(chatId, "🔑 Masukkan Kode Petugas:");
    }

    if (text === "🧑‍💼Login Manajemen") {
        loginState[chatId] = { step: "login_admin" };
        return kirimPesan(chatId, "🔑 Masukkan Kode Manajemen:");
    }

    if (loginState[chatId]?.step === "login_petugas") {
        if (text === kodePetugas) {
            verifiedUsers[chatId] = { name: msg.from.first_name, role: "petugas" };
            loginState[chatId] = null;
            await kirimPesan(chatId, "✅ Login sebagai Petugas");
            return showMenu(chatId);
        } else {
            return kirimPesan(chatId, "❌ Kode salah\n\nSilakan pilih role kembali:", {
                reply_markup: { keyboard: [["👷Login Petugas"], ["🧑‍💼Login Manajemen"]], resize_keyboard: true }
            });
        }
    }

    if (loginState[chatId]?.step === "login_admin") {
        if (text === kodeManajemen) {
            verifiedUsers[chatId] = { name: msg.from.first_name, role: "admin" };
            loginState[chatId] = null;
            await kirimPesan(chatId, "✅ Login sebagai Manajemen");
            return showMenu(chatId);
        } else {
            return kirimPesan(chatId, "❌ Kode salah\n\nSilakan pilih role kembali:", {
                reply_markup: { keyboard: [["👷Login Petugas"], ["🧑‍💼Login Manajemen"]], resize_keyboard: true }
            });
        }
    }

    if (!verifiedUsers[chatId]) return kirimPesan(chatId, "❌ Silakan /start dulu");

    if (text === "🚪 Log out") {
        delete verifiedUsers[chatId];
        delete userState[chatId];
        loginState[chatId] = { step: "pilih_role" };
        return kirimPesan(chatId, "Log out berhasil", {
            reply_markup: { keyboard: [["👷Login Petugas"], ["🧑‍💼Login Manajemen"]], resize_keyboard: true }
        });
    }

    // ==================
    // INPUT DETAIL PELANGGARAN
    // ==================
    if (userState[chatId]?.mode === "input_detail_pelanggaran") {
        userState[chatId].detailPelanggaran = text;
        userState[chatId].mode = "tunggu_lokasi_pelanggaran";
        return kirimPesan(chatId, `✅ Detail tersimpan\n\n📍 Tekan tombol untuk kirim lokasi:`, {
            reply_markup: {
                keyboard: [[{ text: "📍 Kirim Lokasi & Selesai", request_location: true }]],
                resize_keyboard: true,
                one_time_keyboard: true
            }
        });
    }

    // ==================
    // INPUT KETERANGAN HUJAN
    // ==================
    if (userState[chatId]?.mode === "input_keterangan_hujan") {
        if (userState[chatId].mediaHujanCount === 0) {
            return kirimPesan(chatId, "⚠️ Belum ada foto/video yang dikirim. Silakan kirim dulu minimal 1 foto atau video.");
        }
        userState[chatId].keteranganHujan = text;
        userState[chatId].mode = "tunggu_lokasi_hujan";
        return kirimPesan(chatId,
`✅ Keterangan tersimpan

📸 Media    : ${userState[chatId].mediaHujanCount} file
📝 Keterangan : ${text}

📍 Tekan tombol di bawah untuk kirim lokasi & selesai:`, {
            reply_markup: {
                keyboard: [[{ text: "📍 Selesai + Kirim Lokasi", request_location: true }]],
                resize_keyboard: true,
                one_time_keyboard: true
            }
        });
    }

    // ==================
    // HISTORY PENYISIRAN
    // ==================
    if (text === "📋 History Penyisiran") {
        let nama = verifiedUsers[chatId].name;
        let dataUser = laporan.filter(x => x.user === nama);
        let dataHujanUser = laporanHujan.filter(x => x.user === nama);

        if (dataUser.length === 0 && dataHujanUser.length === 0) return kirimPesan(chatId, "❌ Belum ada history penyisiran");

        let rekapSegment = {};
        dataUser.forEach(d => {
            if (!rekapSegment[d.segment]) rekapSegment[d.segment] = [];
            rekapSegment[d.segment].push(d.kp);
        });

        let rekapBersemak = {};
        dataUser.filter(d => d.analisaROW === "Bersemak").forEach(d => {
            if (!rekapBersemak[d.segment]) rekapBersemak[d.segment] = [];
            rekapBersemak[d.segment].push(d.kp);
        });

        let rekapPelanggaran = dataUser.filter(d => d.analisaROW === "Ada Pelanggaran");

        let hasil = `📋 HISTORY PENYISIRAN\n👤 ${nama}\n\n`;
        hasil += `━━━━━━━━━━━━━━━━━━━━\n📍 REKAP PENYISIRAN\n━━━━━━━━━━━━━━━━━━━━\n`;

        let totalTitikAll = 0;
        let totalMeterAll = 0;

        Object.keys(rekapSegment).forEach(seg => {
            let kpList = rekapSegment[seg];
            let titik = kpList.length;
            totalTitikAll += titik;

            let sesi = buatSesi(kpList);
            let meterSegTotal = hitungMeterSesi(sesi);
            totalMeterAll += meterSegTotal;

            hasil += `\n🗺 ${seg}\n`;
            hasil += `   📊 Total : ${titik} titik / ${meterSegTotal} meter\n`;
            sesi.forEach(s => {
                let meterSesi = s.akhir - s.awal;
                hasil += `   • KP ${formatKP(s.awal)} - ${formatKP(s.akhir)} (${s.titik} titik / ${meterSesi} meter)\n`;
            });
        });

        hasil += `\n━━━━━━━━━━━━━━━━━━━━\n`;
        hasil += `📊 TOTAL : ${totalTitikAll} titik / ${totalMeterAll} meter\n`;

        hasil += `\n━━━━━━━━━━━━━━━━━━━━\n🌿 BERSEMAK\n━━━━━━━━━━━━━━━━━━━━\n`;
        if (Object.keys(rekapBersemak).length === 0) {
            hasil += `   Tidak ada data bersemak\n`;
        } else {
            Object.keys(rekapBersemak).forEach(seg => {
                let sesi = buatSesi(rekapBersemak[seg]);
                hasil += `\n🗺 ${seg}\n`;
                sesi.forEach(s => {
                    let meterSesi = s.akhir - s.awal;
                    hasil += `   • KP ${formatKP(s.awal)} - ${formatKP(s.akhir)} (${s.titik} titik / ${meterSesi} meter)\n`;
                });
            });
        }

        hasil += `\n━━━━━━━━━━━━━━━━━━━━\n🚧 FINDING / PELANGGARAN\n━━━━━━━━━━━━━━━━━━━━\n`;
        if (rekapPelanggaran.length === 0) {
            hasil += `   Tidak ada data pelanggaran\n`;
        } else {
            rekapPelanggaran.forEach((d, i) => {
                hasil += `\n${i+1}. 📍 ${d.segment} - KP ${d.kp}\n`;
                hasil += `   🚨 Jenis  : ${d.jenisPelanggaran || "-"}\n`;
                hasil += `   📝 Detail : ${d.detailPelanggaran || "-"}\n`;
            });
        }

        hasil += `\n━━━━━━━━━━━━━━━━━━━━\n🌧 INFORMASI HUJAN\n━━━━━━━━━━━━━━━━━━━━\n`;
        if (dataHujanUser.length === 0) {
            hasil += `   Tidak ada laporan hujan\n`;
        } else {
            dataHujanUser.forEach((d, i) => {
                let koordinat = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "Tidak ada";
                hasil += `\n${i+1}. 📅 ${d.waktu}\n`;
                hasil += `   🌐 Lokasi   : ${koordinat}\n`;
                hasil += `   📝 Keterangan : ${d.keterangan || "-"}\n`;
                hasil += `   📸 Media    : ${d.mediaCount || 0} file\n`;
            });
        }

        return kirimPesan(chatId, hasil);
    }

    // ==================
    // REAL-TIME PENYISIRAN ALL AREA (PETUGAS)
    // ==================
    if (text === "🗺 Real-Time Penyisiran All Area") {
        if (laporan.length === 0) return kirimPesan(chatId, "❌ Belum ada data penyisiran dari petugas manapun");

        // Rekap per petugas per segment
        let rekapPetugas = {};
        laporan.forEach(d => {
            let key = `${d.user}||${d.segment}`;
            if (!rekapPetugas[key]) {
                rekapPetugas[key] = {
                    user: d.user,
                    segment: d.segment,
                    kpList: [],
                    waktuTerakhir: d.waktu
                };
            }
            rekapPetugas[key].kpList.push(d.kp);
            // Simpan waktu paling akhir
            rekapPetugas[key].waktuTerakhir = d.waktu;
        });

        let hasil = `🗺 REAL-TIME PENYISIRAN ALL AREA\n`;
        hasil += `📅 ${new Date().toLocaleString('id-ID')}\n`;
        hasil += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

        hasil += `👷 AKTIVITAS PENYISIRAN PETUGAS\n`;
        hasil += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;

        // Kelompokkan per petugas
        let perPetugas = {};
        Object.values(rekapPetugas).forEach(item => {
            if (!perPetugas[item.user]) perPetugas[item.user] = [];
            perPetugas[item.user].push(item);
        });

        let noPetugas = 1;
        Object.keys(perPetugas).forEach(nama => {
            hasil += `\n${noPetugas}. 👤 ${nama}\n`;
            perPetugas[nama].forEach(item => {
                let sesi = buatSesi(item.kpList);
                sesi.forEach(s => {
                    hasil += `   🗺 ${item.segment}\n`;
                    hasil += `   📍 KP ${formatKP(s.awal)} → ${formatKP(s.akhir)} (${s.titik} titik)\n`;
                    hasil += `   🕐 Terakhir : ${item.waktuTerakhir}\n`;
                });
            });
            noPetugas++;
        });

        // Rekap finding/pelanggaran seluruh area
        let dataFinding = laporan.filter(x => x.analisaROW === "Ada Pelanggaran");

        hasil += `\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        hasil += `🚨 FINDING/PELANGGARAN SELURUH AREA\n`;
        hasil += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;

        if (dataFinding.length === 0) {
            hasil += `\n   ✅ Tidak ada finding/pelanggaran\n`;
        } else {
            dataFinding.forEach((d, i) => {
                hasil += `\n${i+1}. 👤 ${d.user}\n`;
                hasil += `   🗺 ${d.segment} — KP ${d.kp}\n`;
                hasil += `   🚨 Jenis  : ${d.jenisPelanggaran || "-"}\n`;
                hasil += `   📝 Detail : ${d.detailPelanggaran || "-"}\n`;
            });
        }

        return kirimPesan(chatId, hasil);
    }

    // ==================
    // PINDAH SEGMENT
    // ==================
    if (text === "🔀 Pindah Segment") {
        userState[chatId] = { segment: null, kp: null, fotoCount: 0, photos: [], koordinat: null };
        return kirimPesan(chatId, "Pilih Segment baru:", {
            reply_markup: {
                keyboard: [
                    ["Segment 1","Segment 2","Segment 3"],
                    ["Segment 4","Segment 5","Segment 6"],
                    ["Segment 7","Segment 8","Segment 9"],
                    ["Segment 10/12","Segment 11A","Segment 11B"]
                ],
                resize_keyboard: true
            }
        });
    }

    // ==================
    // INFORMASI HUJAN
    // ==================
    if (text === "🌧 Report Hujan") {
        if (verifiedUsers[chatId].role !== "petugas") return kirimPesan(chatId, "❌ Akses ditolak");
        userState[chatId] = {
            ...userState[chatId],
            modeHujan: "tunggu_media_hujan",
            mediaHujan: [],
            mediaHujanCount: 0,
            keteranganHujan: null,
            koordinatHujan: null
        };
        return kirimPesan(chatId,
`🌧 LAPORAN INFORMASI HUJAN

📷 Kirim Bukti foto atau video kondisi hujan.

Setelah semua media terkirim, ketik keterangan detail kondisi hujan.`);
    }

    // ==================
    // DASHBOARD
    // ==================
    if (text === "📊 Dashboard Live Penyisiran Jalur ROW") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");

        let totalMeter = 0;
        let rekapSemuaKP = {};
        laporan.forEach(d => {
            let key = `${d.user}_${d.segment}`;
            if (!rekapSemuaKP[key]) rekapSemuaKP[key] = [];
            rekapSemuaKP[key].push(d.kp);
        });
        Object.values(rekapSemuaKP).forEach(kpList => {
            let sesi = buatSesi(kpList);
            totalMeter += hitungMeterSesi(sesi);
        });

        let segAman = {};
        laporan.filter(x => x.analisaROW === "Tidak Ada Temuan").forEach(d => {
            if (!segAman[d.segment]) segAman[d.segment] = [];
            segAman[d.segment].push(d.kp);
        });
        let infoAman = Object.keys(segAman).length === 0 ? "   Tidak ada data"
            : Object.keys(segAman).map(s => `   • ${s} (${segAman[s].length} titik)`).join("\n");

        let segBersemak = {};
        laporan.filter(x => x.analisaROW === "Bersemak").forEach(d => {
            if (!segBersemak[d.segment]) segBersemak[d.segment] = [];
            segBersemak[d.segment].push(d.kp);
        });
        let infoBersemak = Object.keys(segBersemak).length === 0 ? "   Tidak ada data"
            : Object.keys(segBersemak).map(s => `   • ${s} (${segBersemak[s].length} titik)`).join("\n");

        let segPelanggaran = {};
        laporan.filter(x => x.analisaROW === "Ada Pelanggaran").forEach(d => {
            if (!segPelanggaran[d.segment]) segPelanggaran[d.segment] = [];
            segPelanggaran[d.segment].push(d.kp);
        });
        let infoPelanggaran = Object.keys(segPelanggaran).length === 0 ? "   Tidak ada data"
            : Object.keys(segPelanggaran).map(s => `   • ${s} - KP ${segPelanggaran[s].join(", ")}`).join("\n");

        let infoHujan = laporanHujan.length === 0 ? "   Tidak ada laporan hujan"
            : laporanHujan.map((d, i) => {
                let koord = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "Tidak ada";
                return `   ${i+1}. 👤 ${d.user}\n      📅 ${d.waktu}\n      🌐 ${koord}\n      📝 ${d.keterangan || "-"}`;
            }).join("\n");

        return kirimPesan(chatId,
`📊 DASHBOARD LIVE PENYISIRAN JALUR PIPA ROW PERTAMINA GAS\nOperation Rokan Area\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📏 Jumlah Total Kilometer Pipa Penyisiran Jalur Pipa Row - Pertagas Ora  = [[ ${totalMeter} meter ]]

\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n✅ Jalur ROW Aman       :
${infoAman}

\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n🌿Jalur Bersemak          :
${infoBersemak}

\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n🚧 Finding/Pelanggaran :
${infoPelanggaran}

\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n🌧 Informasi Hujan        :
${infoHujan}`);
    }

    // ==================
    // DATA PETUGAS
    // ==================
    if (text === "👷 Data Petugas Jalur ROW") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");
        if (laporan.length === 0 && laporanHujan.length === 0) return kirimPesan(chatId, "Belum ada data petugas");

        let rekap = {};
        laporan.forEach(d => {
            if (!rekap[d.user]) rekap[d.user] = { totalTitik: 0, segments: {}, hujan: [] };
            rekap[d.user].totalTitik++;
            if (!rekap[d.user].segments[d.segment]) rekap[d.user].segments[d.segment] = [];
            rekap[d.user].segments[d.segment].push(d.kp);
        });

        laporanHujan.forEach(d => {
            if (!rekap[d.user]) rekap[d.user] = { totalTitik: 0, segments: {}, hujan: [] };
            rekap[d.user].hujan.push(d);
        });

        let hasil = "👷 Data Petugas Jalur ROW\n\n";
        Object.keys(rekap).forEach(nama => {
            let data = rekap[nama];

            let totalMeterPetugas = 0;
            Object.values(data.segments).forEach(kpList => {
                let sesi = buatSesi(kpList);
                totalMeterPetugas += hitungMeterSesi(sesi);
            });

            hasil += `👤 ${nama}\n`;
            hasil += `📏 Total KP : ${totalMeterPetugas} meter\n`;
            hasil += `📍 Titik    : ${data.totalTitik} titik\n`;
            hasil += `🌧 Lap. Hujan : ${data.hujan.length} laporan\n\n`;

            Object.keys(data.segments).forEach(seg => {
                let sesi = buatSesi(data.segments[seg]);
                let meterSeg = hitungMeterSesi(sesi);
                hasil += `🗺 ${seg} (${meterSeg} meter) :\n`;
                sesi.forEach(s => {
                    let meterSesi = s.akhir - s.awal;
                    hasil += `   • ${s.titik} titik / KP ${formatKP(s.awal)} - ${formatKP(s.akhir)} (${meterSesi} meter)\n`;
                });
                hasil += "\n";
            });

            if (data.hujan.length > 0) {
                hasil += `🌧 Laporan Hujan :\n`;
                data.hujan.forEach((h, i) => {
                    hasil += `   ${i+1}. ${h.waktu} — ${h.keterangan || "-"}\n`;
                });
                hasil += "\n";
            }

            hasil += "─────────────────\n\n";
        });

        return kirimPesan(chatId, hasil);
    }

    // ==================
    // DATA BERSEMAK
    // ==================
    if (text === "🌿 Data Jalur ROW Bersemak") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");

        let data = laporan.filter(x => x.analisaROW === "Bersemak");
        if (data.length === 0) return kirimPesan(chatId, "Tidak ada data bersemak");

        let segBersemak = {};
        data.forEach(d => {
            if (!segBersemak[d.segment]) segBersemak[d.segment] = [];
            segBersemak[d.segment].push(d.kp);
        });

        let hasil = "🌿 Data Jalur ROW Bersemak\n\n";
        Object.keys(segBersemak).forEach(seg => {
            let sesi = buatSesi(segBersemak[seg]);
            let meter = hitungMeterSesi(sesi);
            hasil += `📍 ${seg}\n   ${segBersemak[seg].length} titik / ${meter} meter\n\n`;
        });

        await kirimPesan(chatId, hasil);

        return kirimPesan(chatId, "Pilih tindakan:", {
            reply_markup: {
                keyboard: [
                    ["📥 Export Data Bersemak"],
                    ["🔙 Kembali Menu Utama"]
                ],
                resize_keyboard: true
            }
        });
    }

    // ==================
    // DATA PELANGGARAN
    // ==================
    if (text === "🚧 Data Finding/Pelanggaran Jalur ROW") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");

        let data = laporan.filter(x => x.analisaROW === "Ada Pelanggaran");
        if (data.length === 0) return kirimPesan(chatId, "Tidak ada data pelanggaran");

        let hasil = "🚧 Data Finding/Pelanggaran Jalur ROW\n\n";
        data.forEach((d, i) => {
            let koordinat = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "Tidak ada";
            hasil += `${i+1}. 👤 ${d.user}\n`;
            hasil += `   📍 ${d.segment} - KP ${d.kp}\n`;
            hasil += `   🚨 Jenis   : ${d.jenisPelanggaran || "-"}\n`;
            hasil += `   📝 Detail  : ${d.detailPelanggaran || "-"}\n`;
            hasil += `   🌐 Koord   : ${koordinat}\n`;
            hasil += `   📅 Waktu   : ${d.waktu}\n\n`;
        });

        await kirimPesan(chatId, hasil);

        return kirimPesan(chatId, "Pilih tindakan:", {
            reply_markup: {
                keyboard: [
                    ["📥 Export Data Finding/Pelanggaran"],
                    ["🔙 Kembali Menu Utama"]
                ],
                resize_keyboard: true
            }
        });
    }

    // ==================
    // KEMBALI MENU UTAMA
    // ==================
    if (text === "🔙 Kembali Menu Utama") {
        return showMenu(chatId);
    }

    // ==================
    // EXPORT EXCEL DATA BERSEMAK (ADMIN)
    // ==================
    if (text === "📥 Export Data Bersemak") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");

        let data = laporan.filter(x => x.analisaROW === "Bersemak");
        if (data.length === 0) return kirimPesan(chatId, "❌ Tidak ada data bersemak untuk di-export");

        await kirimPesan(chatId, `⏳ Membuat file Excel Data Bersemak...\nTotal: ${data.length} titik\n\nMohon tunggu...`);

        try {
            const filePath = await buatExcelBersemak();
            await kirimDokumen(chatId, filePath);
            await kirimPesan(chatId, `✅ File Excel Data Bersemak berhasil dikirim!\n📄 data_bersemak_${getTanggal()}.xlsx`);
            setTimeout(() => { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); }, 15000);
        } catch (err) {
            console.error("Export Bersemak error:", err);
            kirimPesan(chatId, "❌ Gagal export Excel Bersemak: " + err.message);
        }

        return showMenu(chatId);
    }

    // ==================
    // EXPORT EXCEL DATA PELANGGARAN (ADMIN)
    // ==================
    if (text === "📥 Export Data Finding/Pelanggaran") {
        if (verifiedUsers[chatId].role !== "admin") return kirimPesan(chatId, "❌ Akses ditolak");

        let data = laporan.filter(x => x.analisaROW === "Ada Pelanggaran");
        if (data.length === 0) return kirimPesan(chatId, "❌ Tidak ada data pelanggaran untuk di-export");

        await kirimPesan(chatId, `⏳ Membuat file Excel Data Pelanggaran...\nTotal: ${data.length} temuan\n\nMohon tunggu...`);

        try {
            const filePath = await buatExcelPelanggaran();
            await kirimDokumen(chatId, filePath);
            await kirimPesan(chatId, `✅ File Excel Data Pelanggaran berhasil dikirim!\n📄 data_pelanggaran_${getTanggal()}.xlsx`);
            setTimeout(() => { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); }, 15000);
        } catch (err) {
            console.error("Export Pelanggaran error:", err);
            kirimPesan(chatId, "❌ Gagal export Excel Pelanggaran: " + err.message);
        }

        return showMenu(chatId);
    }

    // ==================
    // DOKUMENTASI
    // ==================
    if (text === "📷 Dokumentasi Penyisiran Jalur ROW") {
        return kirimPesan(chatId, "Pilih Segment:", {
            reply_markup: {
                keyboard: [
                    ["Segment 1","Segment 2","Segment 3"],
                    ["Segment 4","Segment 5","Segment 6"],
                    ["Segment 7","Segment 8","Segment 9"],
                    ["Segment 10/12","Segment 11A","Segment 11B"]
                ],
                resize_keyboard: true
            }
        });
    }

    if (text.startsWith("Segment")) {
        userState[chatId] = { segment: text, kp: null, fotoCount: 0, photos: [], koordinat: null };
        return kirimPesan(chatId, `✅ Segment: ${text}\n\nMasukkan KP Awal\nContoh: 0+000`);
    }

    if (text.includes("+") && userState[chatId]?.kp === null) {
        if (!userState[chatId]) return kirimPesan(chatId, "❌ Pilih segment dulu");
        userState[chatId].kp = parseKP(text.trim());
        userState[chatId].fotoCount = 0;
        userState[chatId].photos = [];
        return kirimPesan(chatId,
`📍 ${userState[chatId].segment}
📏 KP ${formatKP(userState[chatId].kp)}

📷 Kirim foto dokumentasi:
- 3 foto = ✅ Jalur ROW Aman
- 4 foto = 🌿 Bersemak
- 5 foto = 🚧 Ada Finding/Pelanggaran`);
    }

    if (["🏗 Bangunan Baru","🌱 Tanaman Baru","⛏ Galian Baru","🔧 Illegal Taping","🛢 Tumpahan Crude","📌 Lainnya"].includes(text)) {
        if (!userState[chatId]) return kirimPesan(chatId, "❌ Tidak ada sesi aktif");
        userState[chatId].jenisPelanggaran = text;
        userState[chatId].mode = "input_detail_pelanggaran";
        return kirimPesan(chatId, `📝 Jenis: ${text}\n\nKetik detail pelanggaran secara lengkap:`);
    }

    if (text === "🔄 Ambil Foto Ulang") {
        if (!userState[chatId] || userState[chatId].kp === null) return kirimPesan(chatId, "❌ Tidak ada sesi aktif");
        userState[chatId].photos.forEach(f => { if (f && fs.existsSync(f)) fs.unlinkSync(f); });
        userState[chatId].fotoCount = 0;
        userState[chatId].photos = [];
        return kirimPesan(chatId,
`🔄 Foto direset\n\n📷 Kirim ulang foto untuk KP ${formatKP(userState[chatId].kp)}\n\n• 3 foto = ✅ Jalur Aman\n• 4 foto = 🌿 Bersemak\n• 5 foto = 🚧 Ada Pelanggaran`);
    }

    if (text === "✅ Selesai (Pelanggaran)") {
        if (!userState[chatId] || userState[chatId].kp === null) return kirimPesan(chatId, "❌ Tidak ada sesi aktif");
        return kirimPesan(chatId, "🚨 Pilih jenis pelanggaran:", {
            reply_markup: {
                keyboard: [
                    ["🏗 Bangunan Baru"],
                    ["🌱 Tanaman Baru"],
                    ["⛏ Galian Baru"],
                    ["🔧 Illegal Taping"],
                    ["🛢 Tumpahan Crude"],
                    ["📌 Lainnya"]
                ],
                resize_keyboard: true
            }
        });
    }

    if (text === "📷 Tambah Foto Temuan" || text === "📷 Tambah Foto Pelanggaran") {
        return kirimPesan(chatId, "📷 Silakan kirim foto berikutnya");
    }

    if (text === "📷 Lanjut Dokumentasi") {
        if (!userState[chatId]) return kirimPesan(chatId, "❌ Tidak ada sesi aktif");
        userState[chatId].fotoCount = 0;
        userState[chatId].photos = [];
        userState[chatId].koordinat = null;
        return kirimPesan(chatId,
`📷 Kirim foto untuk KP ${formatKP(userState[chatId].kp)}\n\n• 3 foto = ✅ Jalur Aman\n• 4 foto = 🌿 Bersemak\n• 5 foto = 🚧 Ada Finding/Pelanggaran`);
    }

    // ==================
    // EXPORT SEMUA — PETUGAS vs ADMIN
    // ==================
    if (text === "📤 Export Penyisiran Jalur ROW") {
        if (laporan.length === 0 && laporanHujan.length === 0) return kirimPesan(chatId, "❌ Belum ada data");

        let role = verifiedUsers[chatId].role;
        let nama = verifiedUsers[chatId].name;

        if (role === "petugas") {
            let dataKu = laporan.filter(d => d.user === nama);
            let dataHujanKu = laporanHujan.filter(d => d.user === nama);
            if (dataKu.length === 0 && dataHujanKu.length === 0) return kirimPesan(chatId, "❌ Belum ada data penyisiranmu");

            await kirimPesan(chatId, `⏳ Membuat file Excel milik ${nama}...`);
            await exportExcelSatuPetugas(chatId, nama);

        } else if (role === "admin") {
            let daftarPetugas = [...new Set([
                ...laporan.map(d => d.user),
                ...laporanHujan.map(d => d.user)
            ])];

            await kirimPesan(chatId,
`⏳ Membuat file Excel untuk ${daftarPetugas.length} petugas:

${daftarPetugas.map((p, i) => `${i+1}. ${p}`).join("\n")}

Mohon tunggu...`);

            await exportExcelSemuaPetugas(chatId);
        }
    }
});

// ==================
// HANDLE FOTO
// ==================
bot.on('photo', async (msg) => {
    const chatId = msg.chat.id;

    // Handle foto untuk laporan hujan
    if (userState[chatId]?.modeHujan === "tunggu_media_hujan") {
        const fileId = msg.photo[msg.photo.length - 1].file_id;
        const filePath = await downloadFoto(fileId);

        if (!filePath) {
            userState[chatId].mediaHujanCount++;
            userState[chatId].mediaHujan.push(null);
            await kirimPesan(chatId, `⚠️ Foto ke-${userState[chatId].mediaHujanCount} gagal tersimpan (tetap dihitung)\n\n✍️Silahkan ketik keterangan Detail. CTH : IJIN SWA, KONDISI HUJAN DERAS:`);
        } else {
            userState[chatId].mediaHujanCount++;
            userState[chatId].mediaHujan.push({ type: 'foto', path: filePath });
            await kirimPesan(chatId, `📸 Foto ke-${userState[chatId].mediaHujanCount} tersimpan ✅\n\n✍️Silahkan ketik keterangan Detail. CTH : IJIN SWA, KONDISI HUJAN DERAS:`);
        }
        userState[chatId].mode = "input_keterangan_hujan";
        return;
    }

    if (!userState[chatId] || userState[chatId].kp === null) {
        return kirimPesan(chatId, "❌ Pilih segment dan masukkan KP dulu");
    }

    if (userState[chatId].mode === "input_detail_pelanggaran" ||
        userState[chatId].mode === "tunggu_lokasi_pelanggaran") {
        return kirimPesan(chatId, "📝 Selesaikan langkah sebelumnya dulu");
    }

    const fileId = msg.photo[msg.photo.length - 1].file_id;
    const filePath = await downloadFoto(fileId);

    if (!filePath) {
        userState[chatId].fotoCount++;
        userState[chatId].photos.push(null);
        await kirimPesan(chatId, `⚠️ Foto ke-${userState[chatId].fotoCount} gagal tersimpan\n📌 Tetap dihitung`);
    } else {
        userState[chatId].fotoCount++;
        userState[chatId].photos.push(filePath);
        await kirimPesan(chatId, `📸 Foto ke-${userState[chatId].fotoCount} tersimpan ✅`);
    }

    const count = userState[chatId].fotoCount;

    if (count < 3) return kirimPesan(chatId, `📌 Butuh minimal 3 foto (${count}/3)`);

    if (count === 3) {
        return kirimPesan(chatId, `Pilih tindakan:`, {
            reply_markup: {
                keyboard: [
                    [{ text: "✅ Selesai + Kirim Lokasi (Jalur Aman)", request_location: true }],
                    [{ text: "📷 Tambah Foto Temuan" }],
                    [{ text: "🔄 Ambil Foto Ulang" }]
                ],
                resize_keyboard: true
            }
        });
    }

    if (count === 4) {
        return kirimPesan(chatId, `Pilih tindakan:`, {
            reply_markup: {
                keyboard: [
                    [{ text: "✅ Selesai + Kirim Lokasi (Bersemak)", request_location: true }],
                    [{ text: "📷 Tambah Foto Pelanggaran" }],
                    [{ text: "🔄 Ambil Foto Ulang" }]
                ],
                resize_keyboard: true
            }
        });
    }

    if (count >= 5) {
        return kirimPesan(chatId, `Pilih tindakan:`, {
            reply_markup: {
                keyboard: [
                    [{ text: "✅ Selesai (Pelanggaran)" }],
                    [{ text: "🔄 Ambil Foto Ulang" }]
                ],
                resize_keyboard: true
            }
        });
    }
});

// ==================
// HANDLE VIDEO
// ==================
bot.on('video', async (msg) => {
    const chatId = msg.chat.id;

    // Handle video untuk laporan hujan
    if (userState[chatId]?.modeHujan === "tunggu_media_hujan" || userState[chatId]?.mode === "input_keterangan_hujan") {
        await kirimPesan(chatId, `⏳ Mengunduh video...`);
        const fileId = msg.video.file_id;
        const filePath = await downloadVideo(fileId);

        if (!filePath) {
            userState[chatId].mediaHujanCount++;
            userState[chatId].mediaHujan.push(null);
            await kirimPesan(chatId, `⚠️ Video ke-${userState[chatId].mediaHujanCount} gagal tersimpan (tetap dihitung)\n\n✍️Silahkan ketik keterangan Detail. CTH : IJIN SWA, KONDISI HUJAN DERAS:`);
        } else {
            userState[chatId].mediaHujanCount++;
            userState[chatId].mediaHujan.push({ type: 'video', path: filePath });
            await kirimPesan(chatId, `🎥 Video ke-${userState[chatId].mediaHujanCount} tersimpan ✅\n\n✍️Silahkan ketik keterangan Detail. CTH : IJIN SWA, KONDISI HUJAN DERAS:`);
        }
        userState[chatId].modeHujan = "tunggu_media_hujan";
        userState[chatId].mode = "input_keterangan_hujan";
        return;
    }
});

// ==================
// PROSES HASIL LAPORAN HUJAN
// ==================
async function prosesHasilHujan(chatId) {
    const waktu = new Date().toLocaleString();
    const koordinat = userState[chatId].koordinatHujan;
    const keterangan = userState[chatId].keteranganHujan || "-";
    const mediaList = userState[chatId].mediaHujan.filter(m => m !== null);

    const data = {
        user: verifiedUsers[chatId].name,
        waktu,
        koordinat: koordinat || null,
        keterangan,
        mediaCount: userState[chatId].mediaHujanCount,
        media: mediaList.map(m => ({ type: m.type, path: m.path }))
    };

    laporanHujan.push(data);
    saveData();

    let infoKoord = koordinat ? `\n🌐 Koordinat : ${koordinat.lat}, ${koordinat.lon}` : "";

    await kirimPesan(chatId,
`🌧 LAPORAN HUJAN TERSIMPAN ✅

👤 Petugas     : ${data.user}
📅 Waktu       : ${data.waktu}${infoKoord}
📸 Media       : ${data.mediaCount} file
📝 Keterangan  : ${data.keterangan}`);

    // Reset state hujan
    userState[chatId].modeHujan = null;
    userState[chatId].mediaHujan = [];
    userState[chatId].mediaHujanCount = 0;
    userState[chatId].keteranganHujan = null;
    userState[chatId].koordinatHujan = null;
    userState[chatId].mode = null;

    return showMenu(chatId);
}

// ==================
// PROSES HASIL FOTO
// ==================
async function prosesHasilFoto(chatId) {
    const count = userState[chatId].fotoCount;
    let analisaROW, keterangan, emoji;

    if (count === 3) {
        analisaROW = "Tidak Ada Temuan"; keterangan = "Jalur Row Terpantau Aman"; emoji = "✅";
    } else if (count === 4) {
        analisaROW = "Bersemak"; keterangan = "Ditemukan semak di jalur, perlu penanganan"; emoji = "🌿";
    } else {
        analisaROW = "Ada Pelanggaran"; keterangan = "Ditemukan pelanggaran di jalur, perlu tindak lanjut segera"; emoji = "🚧";
    }

    const koordinat = userState[chatId].koordinat;
    const waktu = new Date().toLocaleString();

    const data = {
        user: verifiedUsers[chatId].name,
        segment: userState[chatId].segment,
        kp: formatKP(userState[chatId].kp),
        analisaROW, keterangan,
        jenisPelanggaran: userState[chatId].jenisPelanggaran || "-",
        detailPelanggaran: userState[chatId].detailPelanggaran || "-",
        koordinat: koordinat || null,
        photos: userState[chatId].photos.filter(p => p !== null),
        waktu
    };

    laporan.push(data);
    saveData();

    let infoKoord = koordinat ? `\n🌐 Koordinat  : ${koordinat.lat}, ${koordinat.lon}` : "";
    let infoTambahan = "";
    if (analisaROW === "Ada Pelanggaran") {
        infoTambahan = `\n🚨 Jenis      : ${data.jenisPelanggaran}\n📝 Detail     : ${data.detailPelanggaran}`;
    }

    await kirimPesan(chatId,
`${emoji} HASIL ANALISIS JALUR

👤 Petugas     : ${data.user}
📍 Segment     : ${data.segment}
📏 KP          : ${data.kp}
📅 Waktu       : ${data.waktu}
📸 Jumlah Foto : ${count} foto${infoKoord}

🌿 Analisa ROW : ${data.analisaROW}${infoTambahan}`);

    userState[chatId].jenisPelanggaran = null;
    userState[chatId].detailPelanggaran = null;
    userState[chatId].koordinat = null;
    userState[chatId].kp += 100;
    userState[chatId].fotoCount = 0;
    userState[chatId].photos = [];

    kirimPesan(chatId, `➡️ Lanjut ke KP ${formatKP(userState[chatId].kp)}`, {
        reply_markup: {
            keyboard: [
                ["📷 Lanjut Dokumentasi"],
                ["📋 History Penyisiran"],
                ["🗺 Real-Time Penyisiran All Area"],
                ["🔀 Pindah Segment"],
                ["🌧 Report Hujan"],
                ["📤 Export Penyisiran Jalur ROW"],
                ["🚪 Log out"]
            ],
            resize_keyboard: true
        }
    });
}

// ==================
// BUAT EXCEL PER PETUGAS (DENGAN SHEET HUJAN)
// ==================
async function buatExcelPetugas(petugas, tanggal) {
    const workbook = new ExcelJS.Workbook();

    // Sheet Laporan Penyisiran
    const sheet = workbook.addWorksheet("Laporan Penyisiran");

    sheet.columns = [
        { header: "No", key: "no", width: 5 },
        { header: "Segment", key: "segment", width: 13 },
        { header: "KP", key: "kp", width: 10 },
        { header: "Analisa ROW", key: "analisaROW", width: 18 },
        { header: "Jenis Pelanggaran", key: "jenisPelanggaran", width: 18 },
        { header: "Detail Pelanggaran", key: "detailPelanggaran", width: 30 },
        { header: "Koordinat", key: "koordinat", width: 25 },
        { header: "Waktu", key: "waktu", width: 18 },
        { header: "Foto 1", key: "foto1", width: 20 },
        { header: "Foto 2", key: "foto2", width: 20 },
        { header: "Foto 3", key: "foto3", width: 20 },
        { header: "Foto 4", key: "foto4", width: 20 },
        { header: "Foto 5", key: "foto5", width: 20 }
    ];

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    headerRow.height = 20;

    let dataPertugas = laporan.filter(d => d.user === petugas);

    for (let idx = 0; idx < dataPertugas.length; idx++) {
        const d = dataPertugas[idx];
        const rowNum = idx + 2;
        const koordinatStr = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "-";

        sheet.addRow({
            no: idx + 1, segment: d.segment, kp: d.kp,
            analisaROW: d.analisaROW,
            jenisPelanggaran: d.jenisPelanggaran || "-",
            detailPelanggaran: d.detailPelanggaran || "-",
            koordinat: koordinatStr, waktu: d.waktu,
            foto1: "", foto2: "", foto3: "", foto4: "", foto5: ""
        });

        sheet.getRow(rowNum).height = 80;

        if (d.analisaROW === "Ada Pelanggaran") {
            sheet.getRow(rowNum).eachCell(cell => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
            });
        } else if (d.analisaROW === "Bersemak") {
            sheet.getRow(rowNum).eachCell(cell => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
            });
        }

        for (let fIdx = 0; fIdx < Math.min(d.photos.length, 5); fIdx++) {
            await embedFoto(workbook, sheet, d.photos[fIdx], 8 + fIdx, rowNum - 1);
        }

        if (d.koordinat) {
            const cell = sheet.getRow(rowNum).getCell('koordinat');
            cell.value = {
                text: `${d.koordinat.lat}, ${d.koordinat.lon}`,
                hyperlink: `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}`
            };
            cell.font = { color: { argb: 'FF0000FF' }, underline: true };
        }
    }

    // Sheet Laporan Hujan
    const sheetHujan = workbook.addWorksheet("🌧 Laporan Hujan");

    sheetHujan.mergeCells('A1:K1');
    sheetHujan.getCell('A1').value = `LAPORAN INFORMASI HUJAN — ${petugas} — ${tanggal}`;
    sheetHujan.getCell('A1').font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
    sheetHujan.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1565C0' } };
    sheetHujan.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    sheetHujan.getRow(1).height = 28;

    sheetHujan.getRow(2).values = [
        'No', 'Waktu', 'Koordinat', 'Google Maps', 'Keterangan',
        'Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Jumlah Media'
    ];
    sheetHujan.getRow(2).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetHujan.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1565C0' } };
    sheetHujan.getRow(2).height = 22;

    sheetHujan.getColumn(1).width = 5;
    sheetHujan.getColumn(2).width = 22;
    sheetHujan.getColumn(3).width = 28;
    sheetHujan.getColumn(4).width = 18;
    sheetHujan.getColumn(5).width = 40;
    sheetHujan.getColumn(6).width = 22;
    sheetHujan.getColumn(7).width = 22;
    sheetHujan.getColumn(8).width = 22;
    sheetHujan.getColumn(9).width = 22;
    sheetHujan.getColumn(10).width = 22;
    sheetHujan.getColumn(11).width = 15;

    let dataHujanPertugas = laporanHujan.filter(d => d.user === petugas);

    for (let idx = 0; idx < dataHujanPertugas.length; idx++) {
        const d = dataHujanPertugas[idx];
        const rowNum = idx + 3;
        const koordinatStr = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "-";

        sheetHujan.getRow(rowNum).values = [
            idx + 1,
            d.waktu,
            koordinatStr,
            '',
            d.keterangan || "-",
            '', '', '', '', '',
            d.mediaCount || 0
        ];

        sheetHujan.getRow(rowNum).height = 85;
        sheetHujan.getRow(rowNum).eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx % 2 === 0 ? 'FFE3F2FD' : 'FFBBDEFB' } };
            cell.alignment = { vertical: 'middle', wrapText: true };
        });

        if (d.koordinat) {
            const cellMaps = sheetHujan.getRow(rowNum).getCell(4);
            cellMaps.value = {
                text: "Buka Maps",
                hyperlink: `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}`
            };
            cellMaps.font = { color: { argb: 'FF0070C0' }, underline: true };
        }

        const fotoHujan = (d.media || []).filter(m => m && m.type === 'foto' && m.path);
        for (let fIdx = 0; fIdx < Math.min(fotoHujan.length, 5); fIdx++) {
            await embedFoto(workbook, sheetHujan, fotoHujan[fIdx].path, 5 + fIdx, rowNum - 1);
        }
    }

    const namaBersih = petugas.replace(/[^a-zA-Z0-9]/g, '_');
    const filePath = path.join(__dirname, `laporan_${namaBersih}_${tanggal}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    return filePath;
}

// ==================
// BUAT EXCEL KHUSUS DATA BERSEMAK
// ==================
async function buatExcelBersemak() {
    const tanggal = getTanggal();
    const workbook = new ExcelJS.Workbook();

    const sheetRingkasan = workbook.addWorksheet("📊 Ringkasan Bersemak");

    sheetRingkasan.mergeCells('A1:F1');
    sheetRingkasan.getCell('A1').value = `LAPORAN DATA JALUR ROW BERSEMAK — ${tanggal}`;
    sheetRingkasan.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
    sheetRingkasan.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    sheetRingkasan.getRow(1).height = 30;

    const data = laporan.filter(x => x.analisaROW === "Bersemak");

    let segBersemak = {};
    data.forEach(d => {
        if (!segBersemak[d.segment]) segBersemak[d.segment] = [];
        segBersemak[d.segment].push(d.kp);
    });

    let petugasBersemak = {};
    data.forEach(d => {
        if (!petugasBersemak[d.user]) petugasBersemak[d.user] = { titik: 0, segments: {} };
        petugasBersemak[d.user].titik++;
        if (!petugasBersemak[d.user].segments[d.segment]) petugasBersemak[d.user].segments[d.segment] = [];
        petugasBersemak[d.user].segments[d.segment].push(d.kp);
    });

    let totalTitik = data.length;
    let totalMeter = 0;
    Object.values(segBersemak).forEach(kpList => {
        totalMeter += hitungMeterSesi(buatSesi(kpList));
    });

    sheetRingkasan.getRow(3).values = ['', 'TOTAL TITIK BERSEMAK', '', totalTitik + ' titik', '', ''];
    sheetRingkasan.getRow(3).font = { bold: true };
    sheetRingkasan.getRow(4).values = ['', 'TOTAL METER BERSEMAK', '', totalMeter + ' meter', '', ''];
    sheetRingkasan.getRow(4).font = { bold: true };

    sheetRingkasan.getRow(6).values = ['', 'SEGMENT', 'JUMLAH TITIK', 'TOTAL METER', 'RENTANG KP', ''];
    sheetRingkasan.getRow(6).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getRow(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
    sheetRingkasan.getRow(6).height = 20;
    sheetRingkasan.getColumn(2).width = 18;
    sheetRingkasan.getColumn(3).width = 18;
    sheetRingkasan.getColumn(4).width = 18;
    sheetRingkasan.getColumn(5).width = 35;

    let rowIdx = 7;
    Object.keys(segBersemak).forEach(seg => {
        let sesi = buatSesi(segBersemak[seg]);
        let meter = hitungMeterSesi(sesi);
        let rentang = sesi.map(s => `KP ${formatKP(s.awal)} - ${formatKP(s.akhir)}`).join(", ");
        sheetRingkasan.getRow(rowIdx).values = ['', seg, segBersemak[seg].length, meter + ' meter', rentang, ''];
        sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
        rowIdx++;
    });

    rowIdx += 1;
    sheetRingkasan.getRow(rowIdx).values = ['', 'PETUGAS', 'JUMLAH TITIK', 'TOTAL METER', '', ''];
    sheetRingkasan.getRow(rowIdx).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
    sheetRingkasan.getRow(rowIdx).height = 20;
    rowIdx++;

    Object.keys(petugasBersemak).forEach(nama => {
        let dp = petugasBersemak[nama];
        let meterPetugas = 0;
        Object.values(dp.segments).forEach(kpList => {
            meterPetugas += hitungMeterSesi(buatSesi(kpList));
        });
        sheetRingkasan.getRow(rowIdx).values = ['', nama, dp.titik, meterPetugas + ' meter', '', ''];
        sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
        rowIdx++;
    });

    const sheetDetail = workbook.addWorksheet("📋 Detail Bersemak");

    sheetDetail.columns = [
        { header: "No", key: "no", width: 5 },
        { header: "Petugas", key: "user", width: 18 },
        { header: "Segment", key: "segment", width: 15 },
        { header: "KP", key: "kp", width: 10 },
        { header: "Koordinat", key: "koordinat", width: 28 },
        { header: "Google Maps", key: "maps", width: 28 },
        { header: "Waktu", key: "waktu", width: 20 },
        { header: "Foto 1", key: "foto1", width: 22 },
        { header: "Foto 2", key: "foto2", width: 22 },
        { header: "Foto 3", key: "foto3", width: 22 },
        { header: "Foto 4", key: "foto4", width: 22 }
    ];

    const hdrDetail = sheetDetail.getRow(1);
    hdrDetail.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    hdrDetail.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } };
    hdrDetail.height = 22;

    for (let idx = 0; idx < data.length; idx++) {
        const d = data[idx];
        const rowNum = idx + 2;
        const koordinatStr = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "-";

        sheetDetail.addRow({
            no: idx + 1, user: d.user, segment: d.segment, kp: d.kp,
            koordinat: koordinatStr,
            maps: d.koordinat ? `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}` : "-",
            waktu: d.waktu,
            foto1: "", foto2: "", foto3: "", foto4: ""
        });

        sheetDetail.getRow(rowNum).height = 85;
        sheetDetail.getRow(rowNum).eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx % 2 === 0 ? 'FFE2EFDA' : 'FFF0FFF0' } };
            cell.alignment = { vertical: 'middle', wrapText: true };
        });

        if (d.koordinat) {
            const cellMaps = sheetDetail.getRow(rowNum).getCell('maps');
            cellMaps.value = {
                text: "Buka Maps",
                hyperlink: `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}`
            };
            cellMaps.font = { color: { argb: 'FF0070C0' }, underline: true };
        }

        for (let fIdx = 0; fIdx < Math.min(d.photos.length, 4); fIdx++) {
            await embedFoto(workbook, sheetDetail, d.photos[fIdx], 7 + fIdx, rowNum - 1);
        }
    }

    const filePath = path.join(__dirname, `data_bersemak_${tanggal}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    return filePath;
}

// ==================
// BUAT EXCEL KHUSUS DATA PELANGGARAN
// ==================
async function buatExcelPelanggaran() {
    const tanggal = getTanggal();
    const workbook = new ExcelJS.Workbook();

    const data = laporan.filter(x => x.analisaROW === "Ada Pelanggaran");

    let segPelanggaran = {};
    data.forEach(d => {
        if (!segPelanggaran[d.segment]) segPelanggaran[d.segment] = [];
        segPelanggaran[d.segment].push(d);
    });

    let petugasPelanggaran = {};
    data.forEach(d => {
        if (!petugasPelanggaran[d.user]) petugasPelanggaran[d.user] = 0;
        petugasPelanggaran[d.user]++;
    });

    let jenisPelanggaran = {};
    data.forEach(d => {
        let jenis = d.jenisPelanggaran || "Lainnya";
        if (!jenisPelanggaran[jenis]) jenisPelanggaran[jenis] = 0;
        jenisPelanggaran[jenis]++;
    });

    const sheetRingkasan = workbook.addWorksheet("📊 Ringkasan Pelanggaran");

    sheetRingkasan.mergeCells('A1:F1');
    sheetRingkasan.getCell('A1').value = `LAPORAN DATA FINDING / PELANGGARAN JALUR ROW — ${tanggal}`;
    sheetRingkasan.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    sheetRingkasan.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    sheetRingkasan.getRow(1).height = 30;
    sheetRingkasan.getColumn(2).width = 22;
    sheetRingkasan.getColumn(3).width = 18;
    sheetRingkasan.getColumn(4).width = 18;
    sheetRingkasan.getColumn(5).width = 35;

    sheetRingkasan.getRow(3).values = ['', 'TOTAL PELANGGARAN', '', data.length + ' temuan', '', ''];
    sheetRingkasan.getRow(3).font = { bold: true };

    sheetRingkasan.getRow(5).values = ['', 'SEGMENT', 'JUMLAH', 'KP', '', ''];
    sheetRingkasan.getRow(5).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getRow(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    sheetRingkasan.getRow(5).height = 20;

    let rowIdx = 6;
    Object.keys(segPelanggaran).forEach(seg => {
        let kpList = segPelanggaran[seg].map(d => d.kp).join(", ");
        sheetRingkasan.getRow(rowIdx).values = ['', seg, segPelanggaran[seg].length, kpList, '', ''];
        sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        rowIdx++;
    });

    rowIdx += 1;
    sheetRingkasan.getRow(rowIdx).values = ['', 'PETUGAS', 'JUMLAH', '', '', ''];
    sheetRingkasan.getRow(rowIdx).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    sheetRingkasan.getRow(rowIdx).height = 20;
    rowIdx++;

    Object.keys(petugasPelanggaran).forEach(nama => {
        sheetRingkasan.getRow(rowIdx).values = ['', nama, petugasPelanggaran[nama], '', '', ''];
        sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        rowIdx++;
    });

    rowIdx += 1;
    sheetRingkasan.getRow(rowIdx).values = ['', 'JENIS PELANGGARAN', 'JUMLAH', '', '', ''];
    sheetRingkasan.getRow(rowIdx).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    sheetRingkasan.getRow(rowIdx).height = 20;
    rowIdx++;

    Object.keys(jenisPelanggaran).forEach(jenis => {
        sheetRingkasan.getRow(rowIdx).values = ['', jenis, jenisPelanggaran[jenis], '', '', ''];
        sheetRingkasan.getRow(rowIdx).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        rowIdx++;
    });

    const sheetDetail = workbook.addWorksheet("📋 Detail Pelanggaran");

    sheetDetail.columns = [
        { header: "No", key: "no", width: 5 },
        { header: "Petugas", key: "user", width: 18 },
        { header: "Segment", key: "segment", width: 15 },
        { header: "KP", key: "kp", width: 10 },
        { header: "Jenis Pelanggaran", key: "jenisPelanggaran", width: 20 },
        { header: "Detail Pelanggaran", key: "detailPelanggaran", width: 35 },
        { header: "Koordinat", key: "koordinat", width: 28 },
        { header: "Google Maps", key: "maps", width: 18 },
        { header: "Waktu", key: "waktu", width: 20 },
        { header: "Foto 1", key: "foto1", width: 22 },
        { header: "Foto 2", key: "foto2", width: 22 },
        { header: "Foto 3", key: "foto3", width: 22 },
        { header: "Foto 4", key: "foto4", width: 22 },
        { header: "Foto 5", key: "foto5", width: 22 }
    ];

    const hdrDetail = sheetDetail.getRow(1);
    hdrDetail.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    hdrDetail.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    hdrDetail.height = 22;

    for (let idx = 0; idx < data.length; idx++) {
        const d = data[idx];
        const rowNum = idx + 2;
        const koordinatStr = d.koordinat ? `${d.koordinat.lat}, ${d.koordinat.lon}` : "-";

        sheetDetail.addRow({
            no: idx + 1, user: d.user, segment: d.segment, kp: d.kp,
            jenisPelanggaran: d.jenisPelanggaran || "-",
            detailPelanggaran: d.detailPelanggaran || "-",
            koordinat: koordinatStr,
            maps: d.koordinat ? `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}` : "-",
            waktu: d.waktu,
            foto1: "", foto2: "", foto3: "", foto4: "", foto5: ""
        });

        sheetDetail.getRow(rowNum).height = 85;
        sheetDetail.getRow(rowNum).eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx % 2 === 0 ? 'FFFFC7CE' : 'FFFFEBEE' } };
            cell.alignment = { vertical: 'middle', wrapText: true };
        });

        if (d.koordinat) {
            const cellMaps = sheetDetail.getRow(rowNum).getCell('maps');
            cellMaps.value = {
                text: "Buka Maps",
                hyperlink: `https://maps.google.com/?q=${d.koordinat.lat},${d.koordinat.lon}`
            };
            cellMaps.font = { color: { argb: 'FF0070C0' }, underline: true };
        }

        for (let fIdx = 0; fIdx < Math.min(d.photos.length, 5); fIdx++) {
            await embedFoto(workbook, sheetDetail, d.photos[fIdx], 9 + fIdx, rowNum - 1);
        }
    }

    const filePath = path.join(__dirname, `data_pelanggaran_${tanggal}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    return filePath;
}

// ==================
// EXPORT UNTUK PETUGAS
// ==================
async function exportExcelSatuPetugas(chatId, nama) {
    try {
        const tanggal = getTanggal();
        const filePath = await buatExcelPetugas(nama, tanggal);
        await kirimDokumen(chatId, filePath);
        await kirimPesan(chatId, `✅ File Excel milikmu terkirim`);
        setTimeout(() => { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); }, 15000);
    } catch (err) {
        console.error("Export error:", err);
        kirimPesan(chatId, "❌ Gagal export Excel: " + err.message);
    }
}

// ==================
// EXPORT UNTUK ADMIN — SEMUA PETUGAS
// ==================
async function exportExcelSemuaPetugas(chatId) {
    try {
        const tanggal = getTanggal();
        let daftarPetugas = [...new Set([
            ...laporan.map(d => d.user),
            ...laporanHujan.map(d => d.user)
        ])];

        for (let petugas of daftarPetugas) {
            const filePath = await buatExcelPetugas(petugas, tanggal);
            await kirimDokumen(chatId, filePath);
            await kirimPesan(chatId, `✅ laporan_${petugas.replace(/[^a-zA-Z0-9]/g,'_')}_${tanggal}.xlsx terkirim`);
            setTimeout(() => { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); }, 15000);
        }

        await kirimPesan(chatId, `📊 Semua ${daftarPetugas.length} file Excel selesai dikirim ✅`);

    } catch (err) {
        console.error("Export error:", err);
        kirimPesan(chatId, "❌ Gagal export Excel: " + err.message);
    }
}

// ==================
// HELPER
// ==================
function parseKP(kp) {
    let [km, meter] = kp.split("+");
    return parseInt(km) * 1000 + parseInt(meter);
}

function formatKP(totalMeter) {
    let km = Math.floor(totalMeter / 1000);
    let meter = totalMeter % 1000;
    return `${km}+${meter.toString().padStart(3,'0')}`;
}

console.log("🔥 BOT PENYISIRAN PRO SIAP 🚀");
