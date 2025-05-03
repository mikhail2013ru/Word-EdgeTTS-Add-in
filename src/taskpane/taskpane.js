/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    // Заменяем стандартный обработчик кнопки
    document.getElementById("run").onclick = synthesizeSpeech;
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }

// Новая функция для синтеза речи
async function synthesizeSpeech() {
  const runButton = document.getElementById("run");
  try {
    // Получаем выбранный голос
    const voiceSelect = document.getElementById("voice-select");
    const selectedVoice = voiceSelect.value;
    // Получаем выделенный текст
    const text = await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      return range.text;
    });

    if (!text.trim()) {
      throw new Error("Выделите текст для синтеза");
    }

    // Блокируем кнопку
    runButton.disabled = true;
    runButton.textContent = "Синтезируется...";

    // Вариант 1: Браузерный синтез (Web Speech API)
    // const utterance = new SpeechSynthesisUtterance(text);
    // const voices = window.speechSynthesis.getVoices();
    // utterance.voice = voices.find(v => v.lang === 'ru-RU') || voices[0];
    // window.speechSynthesis.speak(utterance);

    // Вариант 2: Через локальный сервер Edge TTS
    
    // const audio = new Audio();
    // audio.src = `http://localhost:5000/synthesize?text=${encodeURIComponent(text)}`;
    // audio.play();    

    // 2. Отправляем на сервер Edge TTS
    const response = await fetch('https://localhost:5000/synthesize', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      body: JSON.stringify({
        text: text,
        voice: selectedVoice
      })
    });

    // 3. Проверяем статус ответа
    if (!response.ok) {
      throw new Error(`Сервер вернул ошибку: ${response.status}`);
    }

    // 4. Воспроизводим аудио
    const blob = await response.blob();
    const audioUrl = URL.createObjectURL(blob);
    const audio = new Audio(audioUrl);
    
    audio.onerror = (e) => {
      throw new Error(`Ошибка воспроизведения: ${e.target.error}`);
    };
    
    audio.play();

    // 5. Уборка
    audio.onended = () => {
      URL.revokeObjectURL(audioUrl);
      runButton.disabled = false;
      runButton.textContent = "Синтезировать речь";
    };

  } catch (error) {
    console.error("Ошибка:", error);

    // Восстанавливаем кнопку
    if (runButton) {
      runButton.disabled = false;
      runButton.textContent = "Синтезировать речь";
    }

    // Показываем уведомление
    showNotification(error.message.includes("Failed to fetch") 
      ? "Сервер TTS недоступен. Запустите локальный сервер." 
      : error.message);
  }
}

// Вспомогательная функция для уведомлений
function showNotification(message) {
  const notification = document.createElement("div");
  notification.className = "notification";
  notification.textContent = message;
  document.body.appendChild(notification);
  setTimeout(() => notification.remove(), 3000);
}
