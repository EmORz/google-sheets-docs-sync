/**
 * СИСТЕМА ЗА УМНА СИНХРОНИЗАЦИЯ (VERSION: ULTIMATE)
 * Използва MD5 хеширане и Batch операции за максимална икономия на ресурси.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Автоматизация')
    .addItem('🚀 Стартирай Синхронизация', 'syncSheetToDocsFinal')
    .addToUi();
}

function isRowActive(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;
  if (typeof value === 'string') {
    const upper = value.toUpperCase().trim();
    return upper === 'ДА' || upper === 'TRUE' || upper === '1';
  }
  return false;
}

function syncSheetToDocsFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;

  // 1. ГРУПОВО ЧЕТЕНЕ (Колони A до G)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  const data = dataRange.getValues();
  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  const folderId = folder.getId();
  const timeZone = Session.getScriptTimeZone();

  let resultsArray = [];     // За колони D и E (URL и Статус)
  let backgroundsArray = []; // За цветове на колона E
  let hashArray = [];        // За колона G (MD5 Хеш)
  let statusCount = { created: 0, updated: 0, unchanged: 0, skipped: 0, error: 0 };

  for (let i = 0; i < data.length; i++) {
    let [name, questions, content, currentUrl, currentStatus, isActive, oldHash] = data[i];
    let newStatus = "";
    let bgColor = "#ffffff";
    let finalUrl = currentUrl;
    let finalHash = oldHash;

    // --- ПРОВЕРКА ЗА ИМЕ ---
    if (!name) {
      resultsArray.push([currentUrl, "Пропуснат (няма име)"]);
      backgroundsArray.push(["#ffffff", "#fbbc04"]);
      hashArray.push([oldHash]);
      statusCount.skipped++;
      continue;
    }

    // --- ПРОВЕРКА ЗА АКТИВНОСТ (Колона F) ---
    if (!isRowActive(isActive)) {
      resultsArray.push([currentUrl, "Изключен"]);
      backgroundsArray.push(["#ffffff", "#f3f3f3"]);
      hashArray.push([oldHash]);
      statusCount.skipped++;
      continue;
    }

    // --- ПОДГОТОВКА НА ТЕКСТ ---
    const questionsText = questions ? questions : "(няма въпроси)";
    const contentText = content ? content : "(няма съдържание)";
    const newTextContent = `УСЛУГА: ${name}\n\nВЪПРОСИ:\n${questionsText}\n\nСЪДЪРЖАНИЕ:\n${contentText}`;
    const newHash = computeMD5(newTextContent);

    // --- ⭐ КЛЮЧОВАТА MD5 ОПТИМИЗАЦИЯ ---
    if (currentUrl && oldHash === newHash) {
      resultsArray.push([currentUrl, "Няма промяна (MD5)"]);
      backgroundsArray.push(["#ffffff", "#ffffff"]);
      hashArray.push([newHash]);
      statusCount.unchanged++;
      continue; // Пропускаме отварянето на документа!
    }

    // --- ОБРАБОТКА НА ДОКУМЕНТА (само при промяна или нов) ---
    try {
      let targetDoc;
      let isNew = false;

      if (currentUrl) {
        try {
          targetDoc = DocumentApp.openByUrl(currentUrl);
          // Местене в правилната папка, ако е извън нея
          const docFile = DriveApp.getFileById(targetDoc.getId());
          if (docFile.getParents().next().getId() !== folderId) {
            docFile.moveTo(folder);
          }
        } catch (e) {
          targetDoc = createNewDoc(name, folder);
          finalUrl = targetDoc.getUrl();
          isNew = true;
        }
      } else {
        targetDoc = createNewDoc(name, folder);
        finalUrl = targetDoc.getUrl();
        isNew = true;
      }

      const body = targetDoc.getBody();
      body.clear();
      
      // Дизайн на документа
      body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph("ВЪПРОСИ:").setBold(true);
      body.appendParagraph(questionsText).setBold(false);
      body.appendHorizontalRule();
      body.appendParagraph("СЪДЪРЖАНИЕ:").setBold(true);
      body.appendParagraph(contentText).setBold(false);

      newStatus = (isNew ? "Създаден" : "Обновен") + " " + Utilities.formatDate(new Date(), timeZone, "HH:mm");
      bgColor = isNew ? "#d9ead3" : "#cfe2f3";
      finalHash = newHash;

      if (isNew) statusCount.created++; else statusCount.updated++;

    } catch (err) {
      newStatus = "ГРЕШКА: " + err.message.slice(0, 50);
      bgColor = "#ea4335";
      statusCount.error++;
      finalHash = oldHash; // Запазваме стария хеш при грешка
    }

    resultsArray.push([finalUrl, newStatus]);
    backgroundsArray.push(["#ffffff", bgColor]);
    hashArray.push([finalHash]);
  }

  // 2. ГРУПОВ ЗАПИС В ТАБЛИЦАТА
  if (resultsArray.length > 0) {
    sheet.getRange(2, 4, resultsArray.length, 2).setValues(resultsArray);
    sheet.getRange(2, 4, backgroundsArray.length, 2).setBackgrounds(backgroundsArray);
    sheet.getRange(2, 7, hashArray.length, 1).setValues(hashArray);
  }

  // 3. ОБОБЩЕН ОТЧЕТ
  const summary = `✅ Синхронизация: ${statusCount.created} нови, ${statusCount.updated} обновени, ${statusCount.unchanged} без промяна, ⏭️ ${statusCount.skipped} пропуснати, ❌ ${statusCount.error} грешки.`;
  sheet.getRange(1, 7).setValue(summary);
  
  // Показване на изскачащо съобщение (опционално)
  SpreadsheetApp.getUi().alert('✅ Синхронизацията завърши!', summary, SpreadsheetApp.getUi().ButtonSet.OK);
}

function createNewDoc(name, folder) {
  const newDoc = DocumentApp.create(name);
  DriveApp.getFileById(newDoc.getId()).moveTo(folder);
  return newDoc;
}

function computeMD5(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
