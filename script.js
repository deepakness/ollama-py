function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🎉')
        .addItem("Fetch Ollama Data", "callOllamaAPI")
        .addToUi();
  }
  
  function callOllamaAPI() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = spreadsheet.getSheetByName('Settings');
  
    // Fetch the settings from Settings Sheet
    var settingsRange = settingsSheet.getRange(2, 2, 5, 1);
    var settingsValues = settingsRange.getValues();
  
    var startRow = Number(settingsValues[0][0]);
    var endRow = Number(settingsValues[1][0]);
    var dataSheet = spreadsheet.getSheetByName(settingsValues[2][0]);
    var promptColumns = settingsValues[3][0].split(',').map(function(item) { return letterToNum(item.trim()); });
    var outputColumns = settingsValues[4][0].split(',').map(function(item) { return letterToNum(item.trim()); });
  
    for (var i = startRow - 1; i < endRow; i++) {
      for (var j = 0; j < promptColumns.length; j++) {
        var promptCell = dataSheet.getRange(i + 1, promptColumns[j]);
        var finalPrompt = promptCell.getValue();
  
        if (!finalPrompt.trim()) {
          continue;
        }
  
        var outputCell = dataSheet.getRange(i + 1, outputColumns[j]);
        
        if (outputCell.getValue() === '') {
          var ollamaData = {
            content: finalPrompt
          },
          ollamaOptions = {
              method: 'post',
              contentType: 'application/json',
              payload: JSON.stringify(ollamaData)
          };
  
          try {
            var ollamaResponse = UrlFetchApp.fetch(`<TUNNEL>/api/chat`, ollamaOptions);
            var ollamaTextResponse = ollamaResponse.getContentText();
            // Remove the leading and trailing quotation marks from the JSON response and trim any leading/trailing whitespace
            var ollamaOutput = ollamaTextResponse.slice(1, -1).trim();
            // If there's a trailing quotation mark left, remove it
            if (ollamaOutput.endsWith('"')) {
                ollamaOutput = ollamaOutput.substring(0, ollamaOutput.length - 1);
            }
            // Replace \n with actual new line characters and \" with "
            var formattedOutput = ollamaOutput.replace(/\\n/g, '\n').replace(/\\"/g, '"');
            outputCell.setValue(formattedOutput);
          } catch(e) {
            console.error('Error calling Ollama API: ' + e.toString());
          }
        }
      }  
    }
  }
  
  function letterToNum(letter) {
    letter = letter.toUpperCase();
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }
  