// =============================================
// Google Apps Script - Social Media Auto Post
// Otomasi konten LinkedIn ke Google Sheet
// Jalankan setiap hari jam 08:00 WIB
// =============================================

const SPREADSHEET_ID = "sheet_id";
const SHEET_NAME = "sheet_name";

// ⚠️ Simpan API key via: Extensions > Apps Script > Project Settings > Script Properties
// Key: ANTHROPIC_API_KEY, Value: sk-ant-...
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("your-key");
}

// -----------------------------------------------
// FUNGSI UTAMA — ini yang diset sebagai trigger
// -----------------------------------------------
function runDailyPost() {
  try {
    Logger.log("Memulai generate konten LinkedIn...");
    
    const today = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    
    // 1. Cek apakah hari ini sudah ada entri
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === today) {
        Logger.log("Konten untuk hari ini sudah ada, skip.");
        return;
      }
    }

    // 2. Generate konten via Claude API
    const post = generateLinkedInPost(today);
    if (!post) {
      Logger.log("Gagal generate konten.");
      return;
    }

    // 3. Masukkan ke Google Sheet
    sheet.appendRow([post.date, post.title, post.content]);
    Logger.log("Berhasil menambahkan baris: " + post.title);

    // 4. (Opsional) Kirim notifikasi email
    sendEmailNotification(post);

  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

// -----------------------------------------------
// Generate konten LinkedIn via Claude API
// -----------------------------------------------
function generateLinkedInPost(date) {
  const apiKey = getApiKey();
  if (!apiKey) {
    Logger.log("API key tidak ditemukan di Script Properties.");
    return null;
  }

  // Daftar topik rotasi agar konten beragam
  const topics = [
    "produktivitas dan manajemen waktu",
    "kepemimpinan dan team building",
    "inovasi dan transformasi digital",
    "pengembangan karier dan personal branding",
    "mindset sukses dan growth mindset",
    "komunikasi efektif di tempat kerja",
    "work-life balance dan kesehatan mental"
  ];
  
  // Pilih topik berdasarkan hari dalam minggu agar tidak berulang
  const dayOfWeek = new Date().getDay(); // 0=Minggu, 6=Sabtu
  const topic = topics[dayOfWeek % topics.length];

  const payload = {
    model: "claude-sonnet-4-20250514",
    max_tokens: 1000,
    system: `Kamu adalah asisten konten LinkedIn profesional.
Tugasmu membuat satu konten LinkedIn per permintaan.

Aturan:
1. Buat judul yang menarik dan relevan.
2. Buat konten 200-300 kata, profesional, bernilai, dan actionable.
3. Gunakan bahasa Indonesia yang profesional namun tidak kaku.
4. Akhiri dengan 5 hashtag relevan.
5. Hindari bahasa terlalu kasual.

Kembalikan HANYA JSON valid (tanpa markdown, tanpa backtick):
{
  "date": "YYYY-MM-DD",
  "title": "Judul post",
  "content": "Isi post lengkap dengan hashtag"
}`,
    messages: [
      {
        role: "user",
        content: `Buat konten LinkedIn tentang: ${topic}. Gunakan tanggal: ${date}`
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    Logger.log("API error: " + JSON.stringify(result.error));
    return null;
  }

  const text = result.content[0].text.replace(/```json|```/g, "").trim();
  return JSON.parse(text);
}

// -----------------------------------------------
// Kirim notifikasi email setelah insert (opsional)
// -----------------------------------------------
function sendEmailNotification(post) {
  const email = Session.getActiveUser().getEmail();
  const subject = `[Auto Post] Konten LinkedIn ${post.date} berhasil dibuat`;
  const body = `Halo,\n\nKonten LinkedIn hari ini telah otomatis ditambahkan ke spreadsheet.\n\n` +
    `Tanggal: ${post.date}\n` +
    `Judul: ${post.title}\n\n` +
    `Preview:\n${post.content.substring(0, 300)}...\n\n` +
    `Lihat spreadsheet: https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}\n\n` +
    `Salam,\nOtomasi Social Media`;
  
  MailApp.sendEmail(email, subject, body);
  Logger.log("Email notifikasi dikirim ke: " + email);
}

// -----------------------------------------------
// Setup trigger otomatis (jalankan sekali saja)
// -----------------------------------------------
function setupDailyTrigger() {
  // Hapus trigger lama jika ada
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "runDailyPost") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Buat trigger baru jam 08:00 WIB (UTC+7 = jam 01:00 UTC)
  ScriptApp.newTrigger("runDailyPost")
    .timeBased()
    .everyDays(1)
    .atHour(1)        // jam 01:00 UTC = 08:00 WIB
    .create();

  Logger.log("Trigger harian jam 08:00 WIB berhasil dibuat.");
}
