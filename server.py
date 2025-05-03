from flask import Flask, request, send_file
from flask_cors import CORS
import edge_tts
import asyncio
import io

app = Flask(__name__)
CORS(app)  # Разрешаем запросы из Office Add-in

@app.route('/synthesize', methods=['POST'])
async def synthesize():
    text = request.json.get('text', '')
    voice = request.json.get('voice', 'ru-RU-SvetlanaNeural')
    
    communicate = edge_tts.Communicate(text, voice)
    
    # Создаем аудио в оперативной памяти
    audio_stream = io.BytesIO()
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            audio_stream.write(chunk["data"])
    
    audio_stream.seek(0)
    return send_file(audio_stream, mimetype='audio/mpeg')

if __name__ == '__main__':
    app.run(port=5000, ssl_context='adhoc')  # HTTPS для Office Add-in