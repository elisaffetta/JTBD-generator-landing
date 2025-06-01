/**
 * JTBD Generator для Google Sheets
 * Описание: Генерирует Jobs To Be Done с помощью OpenRouter API и Gemini 2.5 Flash
 * Автор: elisaffetta
 * Контакты: https://t.me/elisaffetta
 * Заказать разработку/внедрение ИИ в рабочие процессы: https://t.me/elisaffetta
 * Версия: 1.0
 * Дата: 2024
 */

// ========== КОНФИГУРАЦИЯ ==========
const CONFIG = {
  OPENROUTER_API_URL: 'https://openrouter.ai/api/v1/chat/completions',
  MODEL: 'google/gemini-2.5-flash-preview-05-20',
  SHEET_NAME: 'JTBD Generator',
  HEADERS: {
    A1: 'Продукт/Ниша/ЦА',
    B1: 'JTBD #1', C1: 'Обоснование #1', D1: 'Google Sheets решение #1', E1: 'Веб-приложение #1',
    F1: 'JTBD #2', G1: 'Обоснование #2', H1: 'Google Sheets решение #2', I1: 'Веб-приложение #2',
    J1: 'JTBD #3', K1: 'Обоснование #3', L1: 'Google Sheets решение #3', M1: 'Веб-приложение #3',
    N1: 'JTBD #4', O1: 'Обоснование #4', P1: 'Google Sheets решение #4', Q1: 'Веб-приложение #4',
    R1: 'JTBD #5', S1: 'Обоснование #5', T1: 'Google Sheets решение #5', U1: 'Веб-приложение #5',
    V1: 'JTBD #6', W1: 'Обоснование #6', X1: 'Google Sheets решение #6', Y1: 'Веб-приложение #6',
    Z1: 'JTBD #7', AA1: 'Обоснование #7', AB1: 'Google Sheets решение #7', AC1: 'Веб-приложение #7',
    AD1: 'JTBD #8', AE1: 'Обоснование #8', AF1: 'Google Sheets решение #8', AG1: 'Веб-приложение #8',
    AH1: 'JTBD #9', AI1: 'Обоснование #9', AJ1: 'Google Sheets решение #9', AK1: 'Веб-приложение #9',
    AL1: 'JTBD #10', AM1: 'Обоснование #10', AN1: 'Google Sheets решение #10', AO1: 'Веб-приложение #10'
  },
  COLORS: {
    HEADER: '#4285F4',
    NEW_GENERATION: '#E8F0FE',
    OLD_GENERATION: '#F8F9FA',
    SEPARATOR: '#DADCE0'
  }
};

// ========== ОСНОВНЫЕ ФУНКЦИИ ==========

/**
 * Создает меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 JTBD Generator')
    .addItem('🚀 Генерировать JTBD', 'showInputDialog')
    .addSeparator()
    .addItem('📊 Настроить таблицу', 'setupSheet')
    .addItem('🔧 Настройки API', 'showApiSettings')
    .addSeparator()
    .addItem('📖 Инструкция', 'showInstructions')
    .addItem('👨‍💻 Автор: elisaffetta', 'showAuthorInfo')
    .addToUi();
}

/**
 * Показывает диалог для ввода информации о продукте
 */
function showInputDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; }
        .container { max-width: 500px; margin: 0 auto; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: 500; color: #1a73e8; }
        textarea { width: 100%; min-height: 120px; padding: 12px; border: 2px solid #dadce0; border-radius: 8px; font-size: 14px; resize: vertical; }
        textarea:focus { outline: none; border-color: #1a73e8; }
        .btn-primary { background: #1a73e8; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 16px; cursor: pointer; width: 100%; }
        .btn-primary:hover { background: #1557b0; }
        .example { background: #f8f9fa; padding: 12px; border-radius: 8px; margin-top: 8px; font-size: 12px; color: #5f6368; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🎯 Генерация JTBD</h2>
        <form>
          <div class="form-group">
            <label for="productInfo">Информация о продукте/нише/целевой аудитории:</label>
            <textarea id="productInfo" placeholder="Опишите ваш продукт, нишу и целевую аудиторию..." required></textarea>
            <div class="example">
              <strong>Пример:</strong> Мобильное приложение для фитнеса, ориентированное на занятых профессионалов 25-40 лет, которые хотят поддерживать форму, но имеют ограниченное время для тренировок.
            </div>
          </div>
          <button type="button" class="btn-primary" onclick="generateJTBD()">🚀 Генерировать JTBD</button>
        </form>
      </div>
      
      <script>
        function generateJTBD() {
          const productInfo = document.getElementById('productInfo').value.trim();
          if (!productInfo) {
            alert('Пожалуйста, введите информацию о продукте!');
            return;
          }
          
          // Закрываем диалог сразу после нажатия кнопки
          google.script.host.close();
          
          google.script.run
            .withSuccessHandler(() => {
              // Диалог уже закрыт, показываем toast уведомление
            })
            .withFailureHandler((error) => {
              // В случае ошибки показываем alert через основное окно
              console.error('Ошибка генерации JTBD:', error.message);
            })
            .processJTBDGeneration(productInfo);
        }
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 JTBD Generator');
}

/**
 * Обрабатывает генерацию JTBD
 */
function processJTBDGeneration(productInfo) {
  try {
    const sheet = getOrCreateSheet();
    const apiKey = getApiKey();
    
    if (!apiKey) {
      throw new Error('API ключ не настроен. Используйте меню "Настройки API"');
    }
    
    // Показываем прогресс
    SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Генерируем JTBD...', 'Прогресс', 30);
    
    // Генерируем JTBD
    const jtbdData = generateJTBDWithAPI(productInfo, apiKey);
    
    // Добавляем данные в таблицу
    addDataToSheet(sheet, productInfo, jtbdData);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ JTBD успешно сгенерированы!', 'Готово', 5);
    
  } catch (error) {
    Logger.log('Ошибка генерации JTBD: ' + error.toString());
    throw new Error('Не удалось сгенерировать JTBD: ' + error.message);
  }
}

/**
 * Генерирует JTBD с помощью OpenRouter API
 */
function generateJTBDWithAPI(productInfo, apiKey) {
  const prompt = `
Ты эксперт по Jobs To Be Done (JTBD) методологии. Твоя задача - проанализировать информацию о продукте/нише/ЦА и сгенерировать 10 реальных, конкретных и практичных JTBD.

Информация о продукте/нише/ЦА: ${productInfo}

Для каждого JTBD предоставь:
1. Формулировку JTBD в формате "Когда я [ситуация], я хочу [мотивация], чтобы [ожидаемый результат]"
2. Краткое обоснование (2-3 предложения) - откуда взялась эта потребность и почему она важна для данной ЦА
3. Концепцию Google Sheets инструмента для решения этой задачи (1-2 предложения)
4. Концепцию веб-приложения/мини-инструмента для решения этой задачи (1-2 предложения, фокус на 1 продукт = 1 JTBD = 1 ЦА)

Формат ответа должен быть JSON:
{
  "jtbds": [
    {
      "jtbd": "текст JTBD",
      "reasoning": "обоснование",
      "sheets_solution": "концепция Google Sheets решения",
      "web_solution": "концепция веб-приложения"
    }
  ]
}

Важно: JTBD должны быть реалистичными, основанными на реальных потребностях пользователей, а не надуманными.
`;

  const payload = {
    model: CONFIG.MODEL,
    messages: [{
      role: 'user',
      content: prompt
    }],
    temperature: 0.7,
    max_tokens: 4000
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': 'https://sheets.google.com',
      'X-Title': 'JTBD Generator'
    },
    payload: JSON.stringify(payload)
  };
  
  const response = UrlFetchApp.fetch(CONFIG.OPENROUTER_API_URL, options);
  const responseData = JSON.parse(response.getContentText());
  
  if (responseData.error) {
    throw new Error(`API Error: ${responseData.error.message}`);
  }
  
  const content = responseData.choices[0].message.content;
  
  try {
    return JSON.parse(content);
  } catch (e) {
    // Если JSON не парсится, пытаемся извлечь JSON из текста
    const jsonMatch = content.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
    throw new Error('Не удалось распарсить ответ от API');
  }
}

/**
 * Добавляет данные в таблицу
 */
function addDataToSheet(sheet, productInfo, jtbdData) {
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  
  // Добавляем разделитель если есть предыдущие данные
  if (lastRow > 1) {
    const separatorRow = newRow;
    sheet.getRange(separatorRow, 1, 1, 41).setBackground(CONFIG.COLORS.SEPARATOR);
    sheet.getRange(separatorRow, 1).setValue('--- Новая генерация ---');
    sheet.setRowHeight(separatorRow, 30);
  }
  
  const dataRow = lastRow > 1 ? newRow + 1 : newRow;
  
  // Подготавливаем данные для вставки с улучшенным разделением
  const rowData = [productInfo];
  
  // Добавляем данные по каждому JTBD с разделителями
  for (let i = 0; i < 10; i++) {
    if (i < jtbdData.jtbds.length) {
      const jtbd = jtbdData.jtbds[i];
      // Добавляем номер JTBD для лучшего разделения
      const jtbdText = `[JTBD ${i+1}] ${jtbd.jtbd || ''}`;
      const reasoningText = `💡 ${jtbd.reasoning || ''}`;
      const sheetsText = `📊 ${jtbd.sheets_solution || ''}`;
      const webText = `🌐 ${jtbd.web_solution || ''}`;
      
      rowData.push(jtbdText, reasoningText, sheetsText, webText);
    } else {
      rowData.push('', '', '', '');
    }
  }
  
  // Вставляем данные
  sheet.getRange(dataRow, 1, 1, rowData.length).setValues([rowData]);
  
  // Применяем форматирование
  formatNewRow(sheet, dataRow);
  
  // Добавляем границы между группами JTBD для лучшего разделения
  addJTBDBorders(sheet, dataRow);
}

/**
 * Форматирует новую строку
 */
function formatNewRow(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 41);
  range.setBackground(CONFIG.COLORS.NEW_GENERATION);
  range.setWrap(true);
  range.setVerticalAlignment('top');
  
  // Устанавливаем высоту строки
  sheet.setRowHeight(row, 120);
  
  // Добавляем жирный шрифт для первого столбца (продукт/ниша)
  sheet.getRange(row, 1).setFontWeight('bold');
}

/**
 * Добавляет границы между группами JTBD для лучшего визуального разделения
 */
function addJTBDBorders(sheet, row) {
  // Добавляем вертикальные границы между группами JTBD
  const borderColumns = [5, 9, 13, 17, 21, 25, 29, 33, 37, 41]; // Конец каждой группы JTBD
  
  borderColumns.forEach(col => {
    if (col <= 41) {
      sheet.getRange(row, col).setBorder(null, null, null, true, null, null, '#4285F4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });
}

/**
 * Получает или создает лист
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    setupSheetHeaders(sheet);
  }
  
  return sheet;
}

/**
 * Настраивает заголовки листа
 */
function setupSheetHeaders(sheet) {
  // Устанавливаем заголовки
  Object.entries(CONFIG.HEADERS).forEach(([cell, value]) => {
    sheet.getRange(cell).setValue(value);
  });
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(1, 1, 1, 41);
  headerRange.setBackground(CONFIG.COLORS.HEADER);
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment('middle');
  
  // Устанавливаем ширину столбцов
  sheet.setColumnWidth(1, 200); // Продукт/Ниша/ЦА
  for (let i = 2; i <= 41; i++) {
    sheet.setColumnWidth(i, 150);
  }
  
  // Замораживаем первую строку
  sheet.setFrozenRows(1);
}

/**
 * Настраивает лист
 */
function setupSheet() {
  try {
    const sheet = getOrCreateSheet();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Лист настроен!', 'Готово', 3);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка настройки листа: ' + error.message);
  }
}

/**
 * Показывает диалог настроек API
 */
function showApiSettings() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; }
        .container { max-width: 500px; margin: 0 auto; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: 500; color: #1a73e8; }
        input { width: 100%; padding: 12px; border: 2px solid #dadce0; border-radius: 8px; font-size: 14px; }
        input:focus { outline: none; border-color: #1a73e8; }
        .btn-primary { background: #1a73e8; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-size: 16px; cursor: pointer; width: 100%; }
        .btn-primary:hover { background: #1557b0; }
        .info { background: #e8f0fe; padding: 12px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; }
        .warning { background: #fef7e0; padding: 12px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🔧 Настройки API</h2>
        <div class="info">
          <strong>ℹ️ Информация:</strong><br>
          Для работы скрипта необходим API ключ от OpenRouter.ai<br>
          Получить ключ можно на сайте: <a href="https://openrouter.ai" target="_blank">openrouter.ai</a>
        </div>
        <div class="warning">
          <strong>⚠️ Безопасность:</strong><br>
          API ключ сохраняется в защищенном хранилище Google Apps Script
        </div>
        <form>
          <div class="form-group">
            <label for="apiKey">OpenRouter API Key:</label>
            <input type="password" id="apiKey" placeholder="Введите ваш API ключ..." required>
          </div>
          <button type="button" class="btn-primary" onclick="saveApiKey()">💾 Сохранить</button>
        </form>
      </div>
      
      <script>
        function saveApiKey() {
          const apiKey = document.getElementById('apiKey').value.trim();
          if (!apiKey) {
            alert('Пожалуйста, введите API ключ!');
            return;
          }
          
          google.script.run
            .withSuccessHandler(() => {
              alert('✅ API ключ сохранен!');
              google.script.host.close();
            })
            .withFailureHandler((error) => {
              alert('❌ Ошибка: ' + error.message);
            })
            .saveApiKey(apiKey);
        }
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🔧 Настройки API');
}

/**
 * Сохраняет API ключ
 */
function saveApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty('OPENROUTER_API_KEY', apiKey);
  } catch (error) {
    throw new Error('Не удалось сохранить API ключ: ' + error.message);
  }
}

/**
 * Получает API ключ
 */
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
}

/**
 * Показывает информацию об авторе
 */
function showAuthorInfo() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; line-height: 1.6; }
        .container { max-width: 500px; margin: 0 auto; text-align: center; }
        .avatar { width: 100px; height: 100px; border-radius: 50%; margin: 20px auto; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); display: flex; align-items: center; justify-content: center; color: white; font-size: 36px; font-weight: bold; }
        .name { font-size: 24px; font-weight: bold; color: #1a73e8; margin: 20px 0; }
        .description { color: #5f6368; margin: 20px 0; }
        .contact-btn { background: #1a73e8; color: white; border: none; padding: 15px 30px; border-radius: 25px; font-size: 16px; cursor: pointer; text-decoration: none; display: inline-block; margin: 10px; }
        .contact-btn:hover { background: #1557b0; }
        .services { background: #f8f9fa; padding: 20px; border-radius: 12px; margin: 20px 0; text-align: left; }
        .service-item { margin: 10px 0; padding: 8px 0; border-bottom: 1px solid #e8eaed; }
        .service-item:last-child { border-bottom: none; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="avatar">E</div>
        <div class="name">elisaffetta</div>
        <div class="description">
          Эксперт по автоматизации и внедрению ИИ в рабочие процессы.<br>
          Создаю мощные инструменты для Google Workspace и веб-приложения.
        </div>
        
        <div class="services">
          <h3>🚀 Услуги разработки:</h3>
          <div class="service-item">📊 Google Apps Script для автоматизации Sheets/Docs</div>
          <div class="service-item">🤖 Интеграция ИИ в рабочие процессы</div>
          <div class="service-item">🌐 Веб-приложения и мини-инструменты</div>
          <div class="service-item">⚡ Автоматизация бизнес-процессов</div>
          <div class="service-item">🔗 API интеграции и чат-боты</div>
        </div>
        
        <a href="https://t.me/elisaffetta" target="_blank" class="contact-btn">
          💬 Написать в Telegram
        </a>
        
        <div style="margin-top: 30px; font-size: 12px; color: #9aa0a6;">
          Этот скрипт создан с ❤️ для автоматизации генерации JTBD
        </div>
      </div>
    </body>
    </html>
  `).setWidth(600).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, '👨‍💻 Автор: elisaffetta');
}

/**
 * Показывает инструкцию
 */
function showInstructions() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; line-height: 1.6; }
        .container { max-width: 700px; margin: 0 auto; }
        h2 { color: #1a73e8; }
        h3 { color: #1557b0; margin-top: 30px; }
        .step { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 10px 0; }
        .highlight { background: #e8f0fe; padding: 10px; border-radius: 8px; margin: 10px 0; }
        code { background: #f1f3f4; padding: 2px 6px; border-radius: 4px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📖 Инструкция по использованию JTBD Generator</h2>
        
        <h3>🚀 Быстрый старт</h3>
        <div class="step">
          <strong>Шаг 1:</strong> Настройте API ключ через меню "Настройки API"
        </div>
        <div class="step">
          <strong>Шаг 2:</strong> Нажмите "Генерировать JTBD" и введите информацию о вашем продукте
        </div>
        <div class="step">
          <strong>Шаг 3:</strong> Получите 10 JTBD с обоснованиями и концепциями решений
        </div>
        
        <h3>💡 Что такое JTBD?</h3>
        <p>Jobs To Be Done (JTBD) - это методология, которая помогает понять, какую "работу" клиенты "нанимают" ваш продукт выполнить.</p>
        
        <div class="highlight">
          <strong>Формат JTBD:</strong> "Когда я [ситуация], я хочу [мотивация], чтобы [ожидаемый результат]"
        </div>
        
        <h3>📊 Структура таблицы</h3>
        <p>Для каждого JTBD генерируется:</p>
        <ul>
          <li><strong>JTBD</strong> - формулировка потребности</li>
          <li><strong>Обоснование</strong> - почему эта потребность важна</li>
          <li><strong>Google Sheets решение</strong> - концепция инструмента в Sheets</li>
          <li><strong>Веб-приложение</strong> - концепция мини-продукта</li>
        </ul>
        
        <h3>🔧 Технические детали</h3>
        <p>Скрипт использует:</p>
        <ul>
          <li>OpenRouter API для доступа к Gemini 2.5 Flash</li>
          <li>Безопасное хранение API ключей</li>
          <li>Автоматическое форматирование таблицы</li>
          <li>Разделение старых и новых генераций</li>
        </ul>
        
        <div class="highlight">
          <strong>💰 Стоимость:</strong> Использование API платное. Проверьте тарифы на openrouter.ai
        </div>
        
        <h3>👨‍💻 Автор и контакты</h3>
        <div class="highlight">
          <strong>Автор:</strong> elisaffetta<br>
          <strong>Telegram:</strong> <a href="https://t.me/elisaffetta" target="_blank">@elisaffetta</a><br>
          <strong>Заказать разработку/внедрение ИИ:</strong> Пишите в Telegram для обсуждения любых задач по автоматизации и внедрению ИИ в рабочие процессы
        </div>
      </div>
    </body>
    </html>
  `).setWidth(800).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, '📖 Инструкция');
}
