Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Fungsi konversi alfabet ke format Italic Unicode
function toUnicodeItalic(text) {
  return text.split('').map(char => {
    const code = char.charCodeAt(0);
    // Huruf besar A-Z (A=0x1D434)
    if (code >= 65 && code <= 90) {
      return String.fromCodePoint(0x1D434 + (code - 65));
    }
    // Huruf kecil a-z (a=0x1D44E)
    if (code >= 97 && code <= 122) {
      if (char === 'h') return 'ℎ';
      return String.fromCodePoint(0x1D44E + (code - 97));
    }
    return char;
  }).join('');
}

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
      // 1. Ambil sel yang sedang dipilih pengguna
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const originalValues = selectedRange.values;
      const updatedValues = [];
      let count = 0;

      // 2. Proses terjemahan langsung di sel yang sama
      for (let r = 0; r < selectedRange.rowCount; r++) {
        const rowData = [];
        for (let c = 0; c < selectedRange.columnCount; c++) {
          const originalCell = originalValues[r][c];

          if (originalCell !== null && originalCell !== undefined && String(originalCell).trim() !== "") {
            const textToTranslate = String(originalCell).trim();
            const translatedRaw = await fetchTranslation(textToTranslate);

            // Miringkan teks terjemahan
            const italicText = toUnicodeItalic(translatedRaw.trim());
            
            // Bungkus kurung menggunakan LTR Mark (\u200E) agar tanda kurung tidak rusak saat berpindah sel
            const safeBracketed = `\u200E(\u200E${italicText}\u200E)\u200E`;

            let finalResult = "";
            if (position === "newline") {
              // Simpan di bawah teks asli dalam 1 sel yang sama
              finalResult = `${textToTranslate}\n${safeBracketed}`;
            } else {
              // Simpan di samping teks asli dalam 1 sel yang sama
              finalResult = `${textToTranslate} ${safeBracketed}`;
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

      // 3. Tulis hasil langsung menimpa ke sel yang sama
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
