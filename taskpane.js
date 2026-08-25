Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Konversi HANYA teks alfabet ke karakter miring Unicode
function toUnicodeItalic(text) {
  return text.split('').map(char => {
    const code = char.charCodeAt(0);
    // Huruf besar A-Z (A=0x1D434)
    if (code >= 65 && code <= 90) {
      return String.fromCodePoint(0x1D434 + (code - 65));
    }
    // Huruf kecil a-z (a=0x1D44E)
    if (code >= 97 && code <= 122) {
      if (char === 'h') return 'ℎ'; // Pengecualian standar Unicode
      return String.fromCodePoint(0x1D44E + (code - 97));
    }
    // Angka, spasi, dan simbol tetap standar
    return char;
  }).join('');
}

// Fetch terjemahan ID -> EN
async function fetchTranslation(text) {
  const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=id|en`;
  const response = await fetch(url);
  const data = await response.json();
  
  if (data && data.responseData && data.responseData.translatedText) {
    return data.responseData.translatedText;
  }
  throw new Error("Gagal mengambil data terjemahan.");
}

async function handleTranslate() {
  const statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "Menerjemahkan...";

  const position = document.querySelector('input[name="position"]:checked').value;

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const originalValues = range.values;
      const updatedValues = [];

      for (let row = 0; row < originalValues.length; row++) {
        const rowValues = [];
        for (let col = 0; col < originalValues[row].length; col++) {
          const originalText = originalValues[row][col];

          if (typeof originalText === "string" && originalText.trim() !== "") {
            const rawTranslation = await fetchTranslation(originalText.trim());
            
            // 1. Miringkan hanya isi teks terjemahannya saja
            const italicText = toUnicodeItalic(rawTranslation.trim());
            
            // 2. Bungkus teks miring secara eksplisit dengan tanda kurung
            const bracketedTranslation = `(${italicText})`;

            let combinedResult = "";
            if (position === "newline") {
              // Posisi Di Bawah (menggunakan line break \n)
              combinedResult = `${originalText.trim()}\n${bracketedTranslation}`;
            } else {
              // Posisi Di Samping
              combinedResult = `${originalText.trim()} ${bracketedTranslation}`;
            }

            rowValues.push(combinedResult);
          } else {
            rowValues.push(originalText);
          }
        }
        updatedValues.push(rowValues);
      }

      // Tulis hasil kembali ke sel
      range.values = updatedValues;
      
      // Aktifkan text wrap otomatis jika memilih posisi Di Bawah
      if (position === "newline") {
        range.format.wrapText = true;
      }

      await context.sync();
      statusDiv.className = "success";
      statusDiv.innerText = "✓ Selesai diterjemahkan!";
    });
  } catch (error) {
    console.error(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Terjadi kendala: " + error.message;
  }
}
