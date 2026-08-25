Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Menggunakan Google Translate API Endpoint (Gratis & Stabil tanpa blokir CORS)
async function fetchTranslation(text) {
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=en&dt=t&q=${encodeURIComponent(text)}`;
  
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Koneksi API gagal (Status: ${response.status})`);
  }
  
  const data = await response.json();
  
  // Mengambil dan menggabungkan hasil terjemahan dari struktur respon Google
  if (data && data[0]) {
    const translatedText = data[0].map(item => item[0]).join('');
    return translatedText;
  }
  
  throw new Error("Gagal membaca hasil terjemahan.");
}

async function handleTranslate() {
  const statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "⏳ Sedang menerjemahkan...";

  const position = document.querySelector('input[name="position"]:checked').value;

  try {
    await Excel.run(async (context) => {
      // 1. Ambil sel yang sedang dipilih user
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const originalValues = selectedRange.values;
      const translatedMatrix = [];
      let totalTranslated = 0;

      // 2. Proses terjemahan tiap sel
      for (let r = 0; r < selectedRange.rowCount; r++) {
        const rowData = [];
        for (let c = 0; c < selectedRange.columnCount; c++) {
          const originalText = originalValues[r][c];

          if (originalText !== null && originalText !== undefined && String(originalText).trim() !== "") {
            const cleanText = String(originalText).trim();
            const translated = await fetchTranslation(cleanText);
            
            // Format tanda kurung standar
            rowData.push(`(${translated.trim()})`);
            totalTranslated++;
          } else {
            rowData.push("");
          }
        }
        translatedMatrix.push(rowData);
      }

      if (totalTranslated === 0) {
        statusDiv.className = "error";
        statusDiv.innerText = "Pilih sel yang ada teksnya terlebih dahulu!";
        return;
      }

      // 3. Tentukan sel tujuan
      let targetRange;
      if (position === "beside") {
        // Kolom sebelah kanan
        targetRange = selectedRange.getOffsetRange(0, selectedRange.columnCount);
      } else {
        // Baris sebelah bawah
        targetRange = selectedRange.getOffsetRange(selectedRange.rowCount, 0);
      }

      // 4. Masukkan teks dan ubah format font menjadi Italic
      targetRange.values = translatedMatrix;
      targetRange.format.font.italic = true;

      await context.sync();
      statusDiv.className = "success";
      statusDiv.innerText = `✓ Berhasil menerjemahkan ${totalTranslated} sel!`;
    });
  } catch (error) {
    console.error("Detail Error:", error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  }
}
