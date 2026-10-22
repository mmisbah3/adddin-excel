Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Request API Google Translate (ID -> EN)
async function fetchTranslation(text) {
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=en&dt=t&q=${encodeURIComponent(text)}`;
  
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`API Gagal (Status: ${response.status})`);
  }
  
  const data = await response.json();
  if (data && data[0]) {
    return data[0].map(item => item[0]).join('');
  }
  throw new Error("Gagal memproses hasil terjemahan.");
}

async function handleTranslate() {
  const statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "⏳ Sedang menerjemahkan...";

  const position = document.querySelector('input[name="position"]:checked').value;

  try {
    await Excel.run(async (context) => {
      // 1. Ambil sel yang dipilih
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const originalValues = selectedRange.values;
      const updatedValues = [];
      let count = 0;

      // 2. Gabungkan teks asli dan hasil terjemahan murni (plain text)
      for (let r = 0; r < selectedRange.rowCount; r++) {
        const rowData = [];
        for (let c = 0; c < selectedRange.columnCount; c++) {
          const originalCell = originalValues[r][c];

          if (originalCell !== null && originalCell !== undefined && String(originalCell).trim() !== "") {
            const textToTranslate = String(originalCell).trim();
            const translatedRaw = await fetchTranslation(textToTranslate);
            const cleanTranslation = translatedRaw.trim();

            let finalResult = "";
            if (position === "newline") {
              // Baris baru di bawah teks asli
              finalResult = `${textToTranslate}\n${cleanTranslation}`;
            } else {
              // Di samping teks asli
              finalResult = `${textToTranslate} ${cleanTranslation}`;
            }

            rowData.push(finalResult);
            count++;
          } else {
            rowData.push(originalCell);
          }
        }
        updatedValues.push(rowData);
      }

      if (count === 0) {
        statusDiv.className = "error";
        statusDiv.innerText = "Pilih sel yang berisi teks terlebih dahulu!";
        return;
      }

      // 3. Masukkan teks ke sel (format sel asli tetap dipertahankan)
      selectedRange.values = updatedValues;

      // Aktifkan Wrap Text otomatis jika memilih opsi baris baru (Di Bawah)
      if (position === "newline") {
        selectedRange.format.wrapText = true;
      }

      await context.sync();
      statusDiv.className = "success";
      statusDiv.innerText = `✓ Berhasil memperbarui ${count} sel!`;
    });
  } catch (error) {
    console.error(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  }
}
