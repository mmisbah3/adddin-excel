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
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();

      let translatedCount = 0;
      let shapeFound = false;

      // 1. Cek apakah ada Shape aktif yang sedang dipilih atau berisi teks
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        const textFrame = shape.textFrame;
        textFrame.load(["hasText", "textRange"]);
      }
      
      try {
        await context.sync();
      } catch (e) {
        // Abaikan jika worksheet tidak mengizinkan akses massal shape
      }

      // 2. Prioritaskan penerjemahan Sel Terpilih (Selected Range)
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const originalValues = selectedRange.values;
      const updatedValues = [];
      let cellCount = 0;

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
              finalResult = `${textToTranslate}\n${cleanTranslation}`;
            } else {
              finalResult = `${textToTranslate} ${cleanTranslation}`;
            }

            rowData.push(finalResult);
            cellCount++;
          } else {
            rowData.push(originalCell);
          }
        }
        updatedValues.push(rowData);
      }

      // Jika ada teks di sel, simpan hasil ke sel
      if (cellCount > 0) {
        selectedRange.values = updatedValues;
        if (position === "newline") {
          selectedRange.format.wrapText = true;
        }
        translatedCount += cellCount;
      }

      // 3. Terjemahkan teks pada Shape di lembar kerja aktif
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        if (shape.textFrame && shape.textFrame.hasText) {
          const textRange = shape.textFrame.textRange;
          textRange.load("text");
          await context.sync();

          const shapeText = textRange.text;
          if (shapeText && shapeText.trim() !== "") {
            const cleanShapeText = shapeText.trim();
            const translatedShape = await fetchTranslation(cleanShapeText);
            const cleanTransResult = translatedShape.trim();

            let finalShapeResult = "";
            if (position === "newline") {
              finalShapeResult = `${cleanShapeText}\n${cleanTransResult}`;
            } else {
              finalShapeResult = `${cleanShapeText} ${cleanTransResult}`;
            }

            textRange.text = finalShapeResult;
            translatedCount++;
            shapeFound = true;
          }
        }
      }

      await context.sync();

      if (translatedCount === 0) {
        statusDiv.className = "error";
        statusDiv.innerText = "Tidak ditemukan teks pada sel atau shape terpilih!";
        return;
      }

      statusDiv.className = "success";
      statusDiv.innerText = `✓ Berhasil memperbarui ${translatedCount} elemen (sel/shape)!`;
    });
  } catch (error) {
    console.error(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  }
}
