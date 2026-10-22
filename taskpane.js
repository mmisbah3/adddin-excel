Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Fungsi pemanggilan HTTP API ramah IE11 / Excel 2016 (XMLHttpRequest)
function fetchTranslation(text, callback) {
  var url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=en&dt=t&q=" + encodeURIComponent(text);
  var xhr = new XMLHttpRequest();
  
  xhr.open("GET", url, true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          var data = JSON.parse(xhr.responseText);
          if (data && data[0]) {
            var translatedText = "";
            for (var i = 0; i < data[0].length; i++) {
              if (data[0][i][0]) {
                translatedText += data[0][i][0];
              }
            }
            callback(null, translatedText);
          } else {
            callback("Gagal membaca hasil terjemahan.", null);
          }
        } catch (e) {
          callback("Format respon tidak valid.", null);
        }
      } else {
        callback("Gagal terhubung ke API (Status: " + xhr.status + ")", null);
      }
    }
  };
  xhr.send();
}

function handleTranslate() {
  var statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "⏳ Sedang menerjemahkan...";

  var radios = document.getElementsByName("position");
  var position = "inline";
  for (var i = 0; i < radios.length; i++) {
    if (radios[i].checked) {
      position = radios[i].value;
      break;
    }
  }

  Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(["values", "rowCount", "columnCount"]);

    return context.sync().then(function () {
      var originalValues = selectedRange.values;
      var promises = [];
      var cellsToUpdate = [];

      // Kumpulkan semua sel yang berisi teks
      for (var r = 0; r < selectedRange.rowCount; r++) {
        for (var c = 0; c < selectedRange.columnCount; c++) {
          var cellVal = originalValues[r][c];
          if (cellVal !== null && cellVal !== undefined && String(cellVal).replace(/^\s+|\s+$/g, '') !== "") {
            cellsToUpdate.push({
              row: r,
              col: c,
              text: String(cellVal).replace(/^\s+|\s+$/g, '')
            });
          }
        }
      }

      if (cellsToUpdate.length === 0) {
        statusDiv.className = "error";
        statusDiv.innerText = "Pilih sel yang berisi teks terlebih dahulu!";
        return;
      }

      // Proses terjemahan secara berurutan
      var updatedValues = JSON.parse(JSON.stringify(originalValues));
      var completed = 0;

      function processNext(index) {
        if (index >= cellsToUpdate.length) {
          // Setelah semua sel selesai diterjemahkan, tulis kembali ke sel
          selectedRange.values = updatedValues;
          if (position === "newline") {
            selectedRange.format.wrapText = true;
          }

          return context.sync().then(function () {
            statusDiv.className = "success";
            statusDiv.innerText = "✓ Berhasil memperbarui " + completed + " sel!";
          });
        }

        var item = cellsToUpdate[index];
        fetchTranslation(item.text, function (err, translated) {
          if (!err && translated) {
            var cleanTrans = translated.replace(/^\s+|\s+$/g, '');
            if (position === "newline") {
              updatedValues[item.row][item.col] = item.text + "\n" + cleanTrans;
            } else {
              updatedValues[item.row][item.col] = item.text + " " + cleanTrans;
            }
            completed++;
          }
          processNext(index + 1);
        });
      }

      processNext(0);
    });
  }).catch(function (error) {
    console.log(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  });
}
