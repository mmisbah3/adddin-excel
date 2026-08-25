Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Konversi karakter Latin ke format miring (Unicode Italic)
function toUnicodeItalic(text) {
  return text.split('').map(char => {
    const code = char.charCodeAt(0);
    // A-Z
    if (code >= 65 && code <= 90) {
      return String.fromCodePoint(0x1D434 + (code - 65));
    }
    // a-z
    if (code >= 97 && code <= 122) {
      if (char === 'h') return 'ℎ';
      return String.fromCodePoint(0x1D44E + (code - 97));
    }
    return char;
  }).join('');
}

// Request terjemahan langsung spesifik ID ke EN
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
  statusDiv.innerText = "Menerjemahkan ke bahasa Inggris...";

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
            const translatedText = await fetchTranslation(originalText.trim());
            const formattedTranslation = `(${toUnicodeItalic(translatedText)})`;

            let combinedResult = "";
            if (position === "newline") {
              combinedResult = `${originalText}\n${formattedTranslation}`;
            } else {
              combinedResult = `${originalText} ${formattedTranslation}`;
            }

            rowValues.push(combinedResult);
          } else {
            rowValues.push(originalText);
          }
        }
        updatedValues.push(rowValues);
      }

      range.values = updatedValues;
      
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