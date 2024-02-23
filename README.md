# From Local LLMs to Google Sheets

Locally run LLMs like `llama2`, `mistral`, `llava`, etc. on your computer and then get outputs inside Google Sheets.

Watch the below video for more information:

[![Local LLMs to Google Sheets](http://img.youtube.com/vi/q6b9MIktSAc/0.jpg)](http://www.youtube.com/watch?v=q6b9MIktSAc)

## Requirements

- Python
- [ollama](https://ollama.com/)
- [ngrok](https://ngrok.com/)
- Gmail account for Google Sheets

## Process

### 1. Install the `ollama` app on your computer

And run the following command after installing:

```
ollama run mistral
```

If you want to use any other model, you can replace `mistral` with other models that [you can find here](https://ollama.com/library).

### 2. Run the Python code

Open a folder in VS Code or in any other code editor, and create a app.py file, and copy-paste the following code in it:

```python
from flask import Flask, request, jsonify
import ollama

app = Flask(__name__)

@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.json
    response = ollama.chat(model='mistral', messages=[{'role': 'user', 'content': data['content']}])
    return jsonify(response['message']['content'])

if __name__ == '__main__':
    app.run(debug=True, port=5001)
```

If you're using any other model than `mistral`, make sure to replace in the above code as well.

### 3. Install and run `ngrok`

Install `ngrok` by following the instruction from the [official website](https://ngrok.com/). You will also have to sign up for a free account, and you'll get a auth token which you can run in the terminal as:

```
ngrok config add-authtoken <AUTH TOKEN>
```

After this, run the following command and it will start listening to the http://localhost:5001 PORT:

```
ngrok http 5001
```

Copy the `https` forwarding URL you get, it will be used in the next step.

### 4. Prepare Google Sheets Apps Script

Now, you just need to copy-paste the following script inside Google Sheets Apps Script:

```javascript
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
```

Replace the `<TUNNEL>` placeholder with the copied `ngrok` URL.

Save the code and run the script, you may need to authenticate for the very first time.

And just like that, you should start seeing your Google Sheets getting outputs from local LLMs running on your computer.