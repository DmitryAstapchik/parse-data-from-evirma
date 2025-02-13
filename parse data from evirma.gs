// Весь код ниже парсит с эвирмы нужные данные и вставляет в заданные места в гугл-таблице. Переменные для настройки ниже.

// ID вашей Google Таблицы (можно взять из URL)
var spreadsheetId = 'spreadsheet id';
// Название листа
var sheetName = 'sheet name';
// Артикул
var article = 'product article';
// Куки
var cookie = "cookie";
// Переменные, в которые записываются данные
var dataFromEvirmaJson, whs, sizes;
// Данные берутся за дату в формате дд.мм.гггг из ячейки С13 

function fetchDataFromEvirma(article) {
  var url = "https://evirma.ru/api/server/article/promo-journal/data2";

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "accept": "application/json, text/plain, */*",
      "accept-language": "ru,en;q=0.9,ru-RU;q=0.8",
      "cookie": cookie,
      "evirma-project": "1342df9c-6ff2-4a42-81e9-5836e52e8a5d",
      "origin": "https://evirma.ru",
      "priority": "u=1, i",
      "referer": "https://evirma.ru/my/pj/232682611",
      "sec-ch-ua": '"Not A(Brand";v="8", "Chromium";v="132", "Google Chrome";v="132"',
      "sec-ch-ua-mobile": "?0",
      "sec-ch-ua-platform": '"Windows"',
      "sec-fetch-dest": "empty",
      "sec-fetch-mode": "cors",
      "sec-fetch-site": "same-origin",
      "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36"
    },
    payload: JSON.stringify({
      "article": article
    }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = response.getContentText();
  return result;
}

function findNodeByValue(obj, searchString) {
  // Если объект - массив, то обрабатываем каждый элемент
  if (Array.isArray(obj)) {
    for (const item of obj) {
      const result = findNodeByValue(item, searchString);
      if (result) {
        return result;
      }
    }
  } else if (typeof obj === 'object' && obj !== null) {
    // Если объект, обходим его ключи
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        // Если найдено строковое значение, проверяем его
        if (obj[key] === searchString) {
          return obj; // Возвращаем весь объект, где найдено значение
        }
        // Рекурсивно продолжаем поиск в свойствах объекта
        const result = findNodeByValue(obj[key], searchString);
        if (result) {
          return result;
        }
      }
    }
  }
  // Если ничего не найдено, возвращаем null
  return null;
}

function parseNameAndValueFromItems(hash, parsedJson) {
  let node = findNodeByValue(parsedJson, hash);
  let parsed = [];
  for (let item in node.items) {
    parsed.push({ name: node.items[item].name, value: node.items[item].value });
  }
  return parsed;
}

function parseDataForDate(json, dateString) {
  let parsedJson = JSON.parse(json);

  let dayString = dateString.substring(0, 2);
  let monthString = dateString.substring(3, 5);
  let yearString = dateString.substring(6, 10);
  let evirmaDateString = `${yearString}-${monthString}-${dayString}`;

  dataFromEvirmaJson = {
    'trafikVsegoPerehodov': [`open-card-count-${evirmaDateString}`, undefined],
    'trafikReklama': [`total-advt-clicks-count-${evirmaDateString}`, undefined],
    'trafikVneshka': [`article-freq-${evirmaDateString}`, undefined],
    'trafikOrganika': [`organic-open-card-count-${evirmaDateString}`, undefined],
    'osnovnieMetrikiKorzini': [`add-to-cart-count-${evirmaDateString}`, undefined],
    'osnovnieMetrikiZakazi': [`order-count-${evirmaDateString}`, undefined],
    'osnovnieMetrikiProdaji': [`sales-count-${evirmaDateString}`, undefined],
    'osnovnieMetrikiZakaziSumma': [`orders-sum-rub-${evirmaDateString}`, undefined],
    'osnovnieMetrikiProdajiSumma': [`sales-sum-rub-${evirmaDateString}`, undefined],
    'osnovnieMetrikiDrrZakazi': [`total-advt-drr-order-${evirmaDateString}`, undefined],
    'osnovnieMetrikiDrrProdaji': [`total-advt-drr-sales-${evirmaDateString}`, undefined],
    'akcii': [`promo-${evirmaDateString}`, undefined],
    'akciiReiting': [`review-rating-${evirmaDateString}`, undefined],
    'akciiOtzivi': [`review-count-${evirmaDateString}`, undefined],
    'cenaSpp': [`spp-${evirmaDateString}`, undefined],
    'cenaDoSpp': [`price-before-spp-${evirmaDateString}`, undefined],
    'konversiiMoyaProcentVikupa': [`my-buyout-percent-${evirmaDateString}`, undefined],
    'konversiiTop12PerehodKorzina': [`top-add-to-cart-conversion-${evirmaDateString}`, undefined],
    'konversiiTop12KorzinaZakaz': [`top-cart-to-order-conversion-${evirmaDateString}`, undefined],
    'konversiiTop12PerehodZakaz': [`top-open-to-order-conversion-${evirmaDateString}`, undefined],
    'konversiiTop12ProcentVikupa': [`top-buyout-percent-${evirmaDateString}`, undefined],
    'konversiiTop12PerehodVikup': [`top-open-to-sale-conversion-${evirmaDateString}`, undefined],
    'konversiiSrednyayaPerehodKorzina': [`avg-add-to-cart-conversion-${evirmaDateString}`, undefined],
    'konversiiSrednyayaKorzinaZakaz': [`avg-cart-to-order-conversion-${evirmaDateString}`, undefined],
    'konversiiSrednyayaPerehodZakaz': [`avg-open-to-order-conversion-${evirmaDateString}`, undefined],
    'konversiiSrednyayaProcentVikupa': [`avg-buyout-percent-${evirmaDateString}`, undefined],
    'konversiiSrednyayaPerehodVikup': [`avg-open-to-sale-conversion-${evirmaDateString}`, undefined],
    'reklamaZatratiAvto': [`auto-advt-cost-price-${evirmaDateString}`, undefined],
    'reklamaProsmotriAvto': [`auto-advt-views-count-${evirmaDateString}`, undefined],
    'reklamaKlikiAvto': [`auto-advt-clicks-count-${evirmaDateString}`, undefined],
    'reklamaZatratiPoisk': [`search-advt-cost-price-${evirmaDateString}`, undefined],
    'reklamaProsmotriPoisk': [`search-advt-views-count-${evirmaDateString}`, undefined],
    'reklamaKlikiPoisk': [`search-advt-clicks-count-${evirmaDateString}`, undefined],
    'kajdiiPosetitelPrinositVZakazah': [`per1-user-order-sum-rub-${evirmaDateString}`, undefined],
    'kajdiiPosetitelPrinositVProdajah': [`per1-user-sales-sum-rub-${evirmaDateString}`, undefined],
    'naOdnuProdajuPerehodovVKartochku': [`per1-sale-visit-count-${evirmaDateString}`, undefined],
    'naOdnuProdajuKorzin': [`per1-sale-basket-count-${evirmaDateString}`, undefined],
    'naOdnuProdajuZakazov': [`per1-sale-order-count-${evirmaDateString}`, undefined],
    'naOdnuProdajuDenegNaReklamu': [`per1-sale-advt-sum-rub-${evirmaDateString}`, undefined]
  };

  Object.keys(dataFromEvirmaJson).forEach(key => {
    let hash = dataFromEvirmaJson[key][0];
    let value = findNodeByValue(parsedJson, hash).value;
    dataFromEvirmaJson[key][1] = value;
  });

  whs = parseNameAndValueFromItems(`whs-${evirmaDateString}`, parsedJson);
  sizes = parseNameAndValueFromItems(`sizes-${evirmaDateString}`, parsedJson);
}

function writeDataRowToSheet(sheet, startCell, dataRow) {
  var range = sheet.getRange(startCell);
  var startRow = range.getRow();
  var startColumn = range.getColumn();
  for (let i = 0; i < dataRow.length; i++) {
    sheet.getRange(startRow, startColumn + i).setValue(dataRow[i]);
  }
}

function writeDataColumnToSheet(sheet, startCell, dataColumn) {
  var range = sheet.getRange(startCell);
  var column = range.getColumn();
  var row = range.getRow();
  for (var i = 0; i < dataColumn.length; i++) {
    sheet.getRange(row + i, column).setValue(dataColumn[i]);
  }
}

function parseEvirmaAndWriteToSheet() {
  var jsonFromEvirma = fetchDataFromEvirma(article);

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  let dateString = sheet.getRange('C13').getDisplayValue();

  parseDataForDate(jsonFromEvirma, dateString);

  var valuesToInsert = [
    dataFromEvirmaJson['trafikVsegoPerehodov'][1],
    dataFromEvirmaJson['trafikReklama'][1],
    dataFromEvirmaJson['trafikVneshka'][1],
    dataFromEvirmaJson['trafikOrganika'][1]
  ];
  writeDataColumnToSheet(sheet, 'C15', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['osnovnieMetrikiKorzini'][1],
    dataFromEvirmaJson['osnovnieMetrikiZakazi'][1],
    dataFromEvirmaJson['osnovnieMetrikiProdaji'][1],
    dataFromEvirmaJson['osnovnieMetrikiZakaziSumma'][1],
    dataFromEvirmaJson['osnovnieMetrikiProdajiSumma'][1],
    `${dataFromEvirmaJson['osnovnieMetrikiDrrZakazi'][1]}%`.replace('.', ','),
    `${dataFromEvirmaJson['osnovnieMetrikiDrrProdaji'][1]}%`.replace('.', ',')
  ];
  writeDataColumnToSheet(sheet, 'C22', valuesToInsert);

  var valuesToInsert = [`${dataFromEvirmaJson['akcii'][1]}`];
  writeDataRowToSheet(sheet, 'C36', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['akciiReiting'][1],
    dataFromEvirmaJson['akciiOtzivi'][1]
  ];
  writeDataRowToSheet(sheet, 'C37', valuesToInsert);

  var valuesToInsert = [
    `${dataFromEvirmaJson['cenaSpp'][1]}%`.replace('.', ','),
    dataFromEvirmaJson['cenaDoSpp'][1]
  ];
  writeDataColumnToSheet(sheet, 'C40', valuesToInsert);

  var valuesToInsert = [`${dataFromEvirmaJson['konversiiMoyaProcentVikupa'][1]}%`.replace('.', ',')];
  writeDataRowToSheet(sheet, 'C54', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['konversiiTop12PerehodKorzina'][1],
    dataFromEvirmaJson['konversiiTop12KorzinaZakaz'][1],
    dataFromEvirmaJson['konversiiTop12PerehodZakaz'][1],
    dataFromEvirmaJson['konversiiTop12ProcentVikupa'][1],
    dataFromEvirmaJson['konversiiTop12PerehodVikup'][1]
  ];
  writeDataColumnToSheet(sheet, 'E51', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['konversiiSrednyayaPerehodKorzina'][1],
    dataFromEvirmaJson['konversiiSrednyayaKorzinaZakaz'][1],
    dataFromEvirmaJson['konversiiSrednyayaPerehodZakaz'][1],
    dataFromEvirmaJson['konversiiSrednyayaProcentVikupa'][1],
    dataFromEvirmaJson['konversiiSrednyayaPerehodVikup'][1]
  ];
  writeDataColumnToSheet(sheet, 'G51', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['reklamaZatratiAvto'][1],
    dataFromEvirmaJson['reklamaZatratiPoisk'][1]
  ];
  writeDataRowToSheet(sheet, 'E60', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['reklamaProsmotriAvto'][1],
    dataFromEvirmaJson['reklamaProsmotriPoisk'][1]
  ];
  writeDataRowToSheet(sheet, 'E62', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['reklamaKlikiAvto'][1],
    dataFromEvirmaJson['reklamaKlikiPoisk'][1]
  ];
  writeDataRowToSheet(sheet, 'E63', valuesToInsert);

  var valuesToInsert = [];
  whs.forEach(wh => valuesToInsert.push([wh.name, wh.value]));
  var startCell = sheet.getRange('C93');
  var startRow = startCell.getRow();
  var startCol = startCell.getColumn();
  var range = sheet.getRange(startRow, startCol, valuesToInsert.length, valuesToInsert[0].length);
  range.setValues(valuesToInsert);

  var valuesToInsert = [];
  sizes.forEach(size => valuesToInsert.push([size.name, size.value]));
  var startCell = sheet.getRange('C109');
  var startRow = startCell.getRow();
  var startCol = startCell.getColumn();
  var range = sheet.getRange(startRow, startCol, valuesToInsert.length, valuesToInsert[0].length);
  range.setValues(valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['kajdiiPosetitelPrinositVZakazah'][1],
    dataFromEvirmaJson['kajdiiPosetitelPrinositVProdajah'][1]
  ];
  writeDataColumnToSheet(sheet, 'C117', valuesToInsert);

  var valuesToInsert = [
    dataFromEvirmaJson['naOdnuProdajuPerehodovVKartochku'][1],
    dataFromEvirmaJson['naOdnuProdajuKorzin'][1],
    dataFromEvirmaJson['naOdnuProdajuZakazov'][1],
    dataFromEvirmaJson['naOdnuProdajuDenegNaReklamu'][1]
  ];
  writeDataColumnToSheet(sheet, 'C121', valuesToInsert);
}
