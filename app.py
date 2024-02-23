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
