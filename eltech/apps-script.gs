// Google Apps Script для сайта ЭлТех
// Лист "Заявки": Дата | Имя | Телефон | Тип техники | Модель | Неисправность | Статус
// Лист "Прайс": Тип техники | Услуга | Цена
const TG_TOKEN = 'ВСТАВЬТЕ_ТОКЕН_БОТА';
const TG_CHAT = 'ВСТАВЬТЕ_CHAT_ID';

function doGet(e) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Прайс');
  const rows = sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0] && r[1]);
  return json({ rows });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents || '{}');
  if (d.website) return json({ ok: false }); // honeypot
  const phone = String(d.phone || '').replace(/\D/g, '');
  if (phone.length !== 11 || !d.name) return json({ ok: false, error: 'bad data' });
  const sh = SpreadsheetApp.getActive().getSheetByName('Заявки');
  const now = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
  sh.appendRow([now, d.name, d.phone, d.type, d.model, d.issue, 'Новая']);
  const text = '🔧 Новая заявка\n' + d.name + ', ' + d.phone + '\n' + d.type + ' ' + (d.model || '') + '\n' + (d.issue || '');
  UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_TOKEN + '/sendMessage', {
    method: 'post', payload: { chat_id: TG_CHAT, text: text }, muteHttpExceptions: true
  });
  return json({ ok: true });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
