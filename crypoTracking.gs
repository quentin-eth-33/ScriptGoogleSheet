// Variables globales à configuer si besoin
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const columnNumber = SHEET.getLastColumn();
SHEET.autoResizeColumn(columnNumber)
const UI = SpreadsheetApp.getUi();
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
let COLUMN_CR, ROW_FIRST_CRYPTO_CR, COLUMN_NFT_REVIEW, ROW_TOTAL_INVESTED_NR, ROW_SC_CR, ROW_REVIEW_CR, ROW_TOTAL_ASSETS_CR, ROW_EUROS_HTI, COLLUMN_HTI;
let LIST_ALL_CRYPTO = [];
const BG_VERT_CLAIR_3 = "#d9ead3";
const BG_ROUGE_CLAIR_3 = "#f4cccc";
const BG_BLEU_CLAIR_3 = "#cfe2f3";
const BG_GRIS_CLAIR_3 = "#f3f3f3";

const CRYPTO_REVIEW = findCellWithValueAllSheet("Bilan Crypto:");
const NFT_REVIEW = findCellWithValueAllSheet("Bilan NFT:");
const HISTORY_TOTAL_INVESTED = findCellWithValueAllSheet("Historique Total Investi:");
const VALIDED_LOSS = findCellWithValueAllSheet("Historique Cryptos Supprimées:");

if (CRYPTO_REVIEW) {
  COLUMN_CR = CRYPTO_REVIEW.column;
  ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue("Bitcoin", COLUMN_CR);
  ROW_SC_CR = getRowIndexInColumnWithValue("Stablecoin", COLUMN_CR);
  ROW_REVIEW_CR = getRowIndexInColumnWithValue("Bilan:", COLUMN_CR);
  ROW_TOTAL_ASSETS_CR = getRowIndexInColumnWithValue("Total Actif:", COLUMN_CR);
  LIST_ALL_CRYPTO = getAllCrypto(COLUMN_CR, ROW_FIRST_CRYPTO_CR);
}

if (NFT_REVIEW) {
  COLUMN_NFT_REVIEW = NFT_REVIEW.column;
  ROW_TOTAL_INVESTED_NR = getRowIndexInColumnWithValue("Total Investi:", COLUMN_NFT_REVIEW);
}

if (HISTORY_TOTAL_INVESTED) {
  COLLUMN_HTI = HISTORY_TOTAL_INVESTED.column;
  ROW_EUROS_HTI = getRowIndexInColumnWithValue("Euros:", COLLUMN_HTI);
}

function onOpen() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const menuItems = [
    { name: 'Add New Transaction', functionName: 'showFormAddNewTransaction' },
    { name: 'Refresh Crypto Price', functionName: 'callCmcApi' },
    { name: 'Add New Crypto', functionName: 'showFormAddNewCrypto' },
    { name: 'Remove Crypto', functionName: 'showRemoveCrypto' },
    { name: 'Add New NFT', functionName: 'showFormAddNewNFT' },
    { name: 'Add Cash In', functionName: 'showFormAddNewCashIn' },
  ];
  spreadsheet.addMenu('Cryptos Function', menuItems);
}

function showFormAddNewTransaction() {
  if (LIST_ALL_CRYPTO.length < 1) {
    console.log("|showFormAddNewTransaction| LIST_ALL_CRYPTO.length < 1");
    UI.alert('Erreur', "|showFormAddNewTransaction| LIST_ALL_CRYPTO.length < 1", UI.ButtonSet.OK);
    return;
  }
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewTransaction')
    .setWidth(600)
    .setHeight(900);

  htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
  UI.showModalDialog(htmlOutput, 'Ajouter une transaction:');
}

function addNewTransaction(selectedAmount, selectedOption, selectedQuantity, selectedDate, isBuy) {
  if (!COLUMN_CR) {
    console.log("|addNewTransaction| !COLUMN_CR");
    UI.alert('Erreur', "|addNewTransaction| !COLUMN_CR", UI.ButtonSet.OK);
    return;
  }
  let background = BG_VERT_CLAIR_3;
  let cell;

  // Format pour des calculs en gs --> x.xxx | Format pour les cellules de la feuille de calcul --> x,xxx (sinon il y a des problèmes lors des calculs)
  let selectedAmountCell = selectedAmount;
  if (selectedAmount.indexOf('.') !== -1) {
    selectedAmountCell = selectedAmount.replace('.', ',');
  } else if (selectedAmount.indexOf(',') !== -1) {
    selectedAmount = selectedAmount.replace(',', '.');
  }

  let selectedQuantityCell = selectedQuantity;
  if (selectedQuantity.indexOf('.') !== -1) {
    selectedQuantityCell = selectedQuantity.replace('.', ',');
  } else if (selectedQuantity.indexOf(',') !== -1) {
    selectedQuantity = selectedQuantity.replace(',', '.');
  }

  console.log("Option sélectionnée : " + selectedOption + " | Date : " + selectedDate + " | Montant : " + selectedAmount + " | Quantité : " + selectedQuantity);

  let selectedOptionIndex = findCellsWithValueAndBackgroundAllSheet(selectedOption, "#d69bff");
  if (selectedOptionIndex) {
    if (isBuy === false) {
      background = BG_ROUGE_CLAIR_3;
    }

    // Recherche de la ligne ou insérer la transaction
    let values = SHEET.getRange((selectedOptionIndex.row + 3), selectedOptionIndex.column, SHEET.getLastRow(), 1).getValues();
    let newTransactionRow = 3;
    for (let j = 0; j < values.length; j++) {
      if (values[j][0] === "") {
        newTransactionRow = j + 3 + selectedOptionIndex.row;
        break;
      }
    }

    let range = SHEET.getRange(newTransactionRow, selectedOptionIndex.column, 1, 4);
    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    range.setBackground(background);
    range.setHorizontalAlignment("right");
    range.setFontWeight("bold");

    cell = range.getCell(1, 1);
    cell.setNumberFormat(getCellFormatNumberDollars(selectedAmount));
    cell.setValue(selectedAmountCell);

    cell = range.getCell(1, 2);
    cell.setNumberFormat(getCellFormatNumber(selectedQuantity));
    cell.setValue(selectedQuantityCell);

    cell = range.getCell(1, 3); // Première cellule de la plage
    let averageBuyingPrice = selectedAmount / selectedQuantity;
    cell.setNumberFormat(getCellFormatNumberDollars(averageBuyingPrice));
    cell.setValue(averageBuyingPrice);


    cell = range.getCell(1, 4); // Première cellule de la plage
    let dateFormat = "dd/MM/YYYY";
    cell.setNumberFormat(dateFormat);
    let dateComponents = selectedDate.split("-");
    let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
    cell.setValue(formattedDate);

    let sellSum = 0;
    let buySum = 0;
    let buyQuantity = 0;
    let sellQuantity = 0;

    range = SHEET.getRange((selectedOptionIndex.row + 3), selectedOptionIndex.column, SHEET.getLastRow(), 4);

    values = range.getValues();
    const backgrounds = range.getBackgrounds();

    for (let i = 0; i < values.length; i++) {
      if (backgrounds[i][0] == BG_VERT_CLAIR_3) {
        buySum += values[i][0];
        buyQuantity += values[i][1];
      } else if (backgrounds[i][0] == BG_ROUGE_CLAIR_3) {
        sellSum += values[i][0];
        sellQuantity += values[i][1];
      } else {
        break;
      }
    }

    console.log("buyQuantity: " + buyQuantity + " | sellQuantity: " + sellQuantity)
    let selectedOptionCR_Row = getRowIndexInColumnWithValue(selectedOption, COLUMN_CR);

    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + 1));
    cell.setNumberFormat(getCellFormatNumberDollars((buySum - sellSum)));
    cell.setValue((buySum - sellSum));

    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + 2));
    cell.setNumberFormat(getCellFormatNumber((buyQuantity - sellQuantity)));
    cell.setValue((buyQuantity - sellQuantity));

    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + 3));
    cell.setNumberFormat(getCellFormatNumberDollars((buySum / buyQuantity)));
    cell.setValue((buySum / buyQuantity));

  } else {
    console.log(selectedOption + "n'a pas été trouvé");
    UI.alert('Erreur', selectedOption + "n'a pas été trouvé", UI.ButtonSet.OK);
  }
}


function callCmcApi() {
  const ui = SpreadsheetApp.getUi();
  if (!ROW_FIRST_CRYPTO_CR || !COLUMN_CR) {
    console.log("|callCmcApi| !ROW_FIRST_CRYPTO_CR || !COLUMN_CR");
    ui.alert('Erreur', "|callCmcApi| !ROW_FIRST_CRYPTO_CR || !COLUMN_CR", UI.ButtonSet.OK);
    return;
  }

  let mapIdCryptoGlobalString = SCRIPT_PROPERTIES.getProperty('mapIdCryptoGlobal');
  refreshEurosPrice();

  // Convertir la chaîne en map
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    mapIdCryptoGlobal[key] = value;
  });

  const startRow = ROW_FIRST_CRYPTO_CR;
  const endRow = SHEET.getLastRow();
  const urlSymbols = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
  const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  const apiKeyList = getApiKeyCmcList();
  let indexTabApiKey = 0;
  const indexColumnCurrentPrice = COLUMN_CR + 4;
  let idsAsString = "";

  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": apiKeyList[indexTabApiKey],
    },
    muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
  };

  let cryptosInfo = {};
  let cryptoId;
  let cryptoPrice;
  let isCallValid;
  for (let row = startRow; row <= endRow; row++) {
    const cryptoName = SHEET.getRange(row, COLUMN_CR).getValue();
    isCallValid = false;
    if (cryptoName === "Stablecoin") {
      break;
    }
    if (!(cryptoName in mapIdCryptoGlobal)) {
      console.log("Pas variable globale: " + cryptoName);
      let response = UrlFetchApp.fetch(urlSymbols, options);
      let jsonData = response.getContentText();

      if (response.getResponseCode() === 200) {
        const json = JSON.parse(jsonData);
        const data = json.data;
        const cryptoData = data.find(item => item.name === cryptoName);

        if (!cryptoData) {
          console.log(`Cryptomonnaie '${cryptoName}' introuvable.`);
          continue;
        }

        cryptoId = cryptoData.id;
        mapIdCryptoGlobalString += `;${cryptoName}:${cryptoId}`;
        SCRIPT_PROPERTIES.setProperty('mapIdCryptoGlobal', mapIdCryptoGlobalString);
        isCallValid = true;
      } else if (response.getResponseCode() === 429) {
        if (indexTabApiKey < apiKeyList.length - 1) {
          indexTabApiKey = indexTabApiKey + 1;
          options.headers["X-CMC_PRO_API_KEY"] = apiKeyList[indexTabApiKey];
          row = row - 1;
        } else {
          console.log("plus d'api key opérationnelle disponible")
        }

      } else {
        console.log(`Erreur lors de la récupération des symboles de cryptomonnaie : ${response.getResponseCode()} - ${jsonData}`);
      }
    }
    else {
      cryptoId = mapIdCryptoGlobal[cryptoName];
      isCallValid = true;
    }

    if (isCallValid) {
      cryptosInfo[cryptoName] = { id: cryptoId, indexRow: row };
    }
  }

  for (const crypto in cryptosInfo) {
    // Vérifier si la propriété "id" existe pour cette cryptomonnaie
    if (cryptosInfo[crypto].hasOwnProperty("id")) {
      if (cryptosInfo[crypto].id !== undefined) {
        idsAsString += cryptosInfo[crypto].id + ",";
      }
    }
  }
  // Supprimer la dernière virgule
  idsAsString = idsAsString.slice(0, -1);

  let quoteResponse = UrlFetchApp.fetch(`${urlQuotes}?id=${idsAsString}`, options);
  let quoteJsonData = quoteResponse.getContentText();

  if (quoteResponse.getResponseCode() === 200) {
    const quoteJson = JSON.parse(quoteJsonData);
    for (const crypto in cryptosInfo) {
      if (cryptosInfo[crypto].hasOwnProperty("id")) {
        if (cryptosInfo[crypto].id !== undefined) {
          cryptoPrice = quoteJson.data[cryptosInfo[crypto].id].quote.USD.price;
          console.log("Crypto: " + crypto + " | Price: " + cryptoPrice)
          if (cryptoPrice != 0) {
            SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setNumberFormat(getCellFormatNumberDollars(cryptoPrice));
            SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setFontWeight("bold");
            SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setHorizontalAlignment("right");
            SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setBackground(BG_GRIS_CLAIR_3);
            SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setValue(cryptoPrice);
          }
        }
      }
    }

  } else if (quoteResponse.getResponseCode() === 429) {
    if (indexTabApiKey < apiKeyList.length - 1) {
      indexTabApiKey = indexTabApiKey + 1;
      options.headers["X-CMC_PRO_API_KEY"] = apiKeyList[indexTabApiKey];
    } else {
      console.log("plus d'api key opérationnelle disponible")
    }
  }
  else {
    console.log(`Erreur lors de la récupération des données de la cryptomonnaie '${cryptoName}' : ${quoteResponse.getResponseCode()} - ${quoteJsonData}`);
  }
}

function showFormAddNewNFT() {
  if (LIST_ALL_CRYPTO.length < 1) {
    console.log("|showFormAddNewNFT| LIST_ALL_CRYPTO.length < 1");
    UI.alert('Erreur', "|showFormAddNewNFT| LIST_ALL_CRYPTO.length < 1", UI.ButtonSet.OK);
    return;
  }
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewNFT')
    .setWidth(600)
    .setHeight(900);

  htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
  UI.showModalDialog(htmlOutput, 'Nouvel Achat NFT:');
}

function addNewNFT(idNftInput, optionSelectSell, amountInput, quantityInputSell, selectedDate) {
  if (!ROW_TOTAL_INVESTED_NR || !COLUMN_NFT_REVIEW) {
    console.log("|addNewNFT| !ROW_TOTAL_INVESTED_NR || !COLUMN_NFT_REVIEW");
    UI.alert('Erreur', "|addNewNFT| !ROW_TOTAL_INVESTED_NR || !COLUMN_NFT_REVIEW", UI.ButtonSet.OK);
    return;
  }
  console.log("Id NFT: " + idNftInput + " | optionSelectSell: " + optionSelectSell + " | amountInput: " + amountInput + " | quantityInputSell: " + quantityInputSell)

  let amountInputCell = amountInput;
  if (amountInput.indexOf('.') !== -1) {
    amountInputCell = amountInput.replace('.', ',');
  } else if (amountInput.indexOf(',') !== -1) {
    amountInput = amountInput.replace(',', '.');
  }

  SHEET.getRange(ROW_TOTAL_INVESTED_NR, COLUMN_NFT_REVIEW, 2, 4).moveTo(SHEET.getRange((ROW_TOTAL_INVESTED_NR + 1), COLUMN_NFT_REVIEW));

  let cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR, COLUMN_NFT_REVIEW, 1, 4);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("right");
  cell.setFontSize(8);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR, COLUMN_NFT_REVIEW);
  cell.setBackground(BG_BLEU_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("left");
  cell.setFontSize(8);
  cell.setFontColor("#4285f4");
  cell.setValue(idNftInput);

  cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR, (COLUMN_NFT_REVIEW + 1));
  cell.setNumberFormat(getCellFormatNumberDollars(amountInput));
  cell.setValue(amountInputCell);

  cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR, (COLUMN_NFT_REVIEW + 2));
  cell.setValue("Non défini");

  cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR, (COLUMN_NFT_REVIEW + 3));
  cell.setNumberFormat("0$");
  cell.setValue(0);

  addNewTransaction(amountInput, optionSelectSell, quantityInputSell, selectedDate, false);
}


function showFormAddNewCrypto() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCrypto')
    .setWidth(300)
    .setHeight(220);

  UI.showModalDialog(htmlOutput, 'Ajouter une nouvelle crypto:');
}

function addNewCrypto(cryptoName) {
  if (!ROW_SC_CR || !COLUMN_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR || !ROW_FIRST_CRYPTO_CR) {
    console.log("|addNewCrypto| !ROW_SC_CR || !COLUMN_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR || !ROW_FIRST_CRYPTO_CR");
    UI.alert('Erreur', "|addNewCrypto| !ROW_SC_CR || !COLUMN_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR || !ROW_FIRST_CRYPTO_CR", UI.ButtonSet.OK);
    return;
  }
  let range, cell;

  const stablecoinTH = findCellsWithValueAndBackgroundAllSheet("Stablecoin", "#d69bff");

  SHEET.insertColumnsBefore((stablecoinTH.column - 1), 5); // -1 -> car sinon ca prend le format de "StableCoin Bincance" (cellule colorées etc))


  range = SHEET.getRange(stablecoinTH.row, stablecoinTH.column, 2, 4)
  range.setBackground("#d69bff");
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  range.setFontSize(20);
  range.setFontColor("#351c75");
  range.setValue(cryptoName);
  range.merge();
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  range = SHEET.getRange((stablecoinTH.row + 2), stablecoinTH.column, 1, 4);
  range.setBackground("#ebcfff");
  range.setFontColor("#9900ff");
  range.setFontWeight("bold");
  range.setFontSize(8);
  range.setHorizontalAlignment("left");
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = range.getCell(1, 1);
  cell.setValue("Montant:")

  cell = range.getCell(1, 2);
  cell.setValue("Quantité:")

  cell = range.getCell(1, 3);
  cell.setValue("Prix:")

  cell = range.getCell(1, 4);
  cell.setValue("Date:")

  SHEET.getRange(ROW_SC_CR, COLUMN_CR, (ROW_REVIEW_CR - ROW_SC_CR + 1), 9).moveTo(SHEET.getRange((ROW_SC_CR + 1), COLUMN_CR));

  cell = SHEET.getRange(ROW_SC_CR, COLUMN_CR);
  cell.setBackground(BG_BLEU_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("left");
  cell.setFontSize(8);
  cell.setFontColor("#4285f4");
  cell.setValue(cryptoName);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 1), 1, 8);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("right");
  cell.setFontSize(8);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 1));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 2));
  cell.setValue(0);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 3));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 4));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 5));
  cell.setFormula("=" + getColumnHeader(COLUMN_CR + 4) + "" + ROW_SC_CR + "*" + getColumnHeader(COLUMN_CR + 2) + "" + ROW_SC_CR);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 6));
  cell.setFormula("=" + getColumnHeader(COLUMN_CR + 5) + "" + ROW_SC_CR + "-" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_SC_CR);
  cell.setNumberFormat('[Color50]+0$;[RED]-0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 7));
  cell.setFormula("=IF(" + getColumnHeader(COLUMN_CR + 3) + "" + ROW_SC_CR + "=0;0;IF(" + getColumnHeader(COLUMN_CR + 4) + "" + ROW_SC_CR + "<" + getColumnHeader(COLUMN_CR + 3) + "" + ROW_SC_CR + ";-(1-(" + getColumnHeader(COLUMN_CR + 4) + "" + ROW_SC_CR + "/" + getColumnHeader(COLUMN_CR + 3) + "" + ROW_SC_CR + "));(" + getColumnHeader(COLUMN_CR + 4) + "" + ROW_SC_CR + "/" + getColumnHeader(COLUMN_CR + 3) + "" + ROW_SC_CR + ")-1))");
  cell.setNumberFormat('[Color50]+0.00%;[Red]-0.00%');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 8));
  cell.setFormula("=" + getColumnHeader(COLUMN_CR + 5) + "" + ROW_SC_CR + "/" + getColumnHeader(COLUMN_CR + 1) + "" + (ROW_SC_CR + 3))
  cell.setNumberFormat('0.00%');

  // Créer la règle de mise en forme conditionnelle
  let plageFormatConditionnelle = SHEET.getRange("" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + 8) + "" + ROW_SC_CR);
  let regleMiseEnForme = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + "<=0")
    .setBackground(BG_VERT_CLAIR_3) // Couleur de fond verte en cas de condition satisfaite
    .setRanges([plageFormatConditionnelle])
    .build();
  SHEET.setConditionalFormatRules([regleMiseEnForme]);

  cell = SHEET.getRange((ROW_TOTAL_ASSETS_CR + 1), (COLUMN_CR + 1));
  cell.setFormula("=SUM(" + getColumnHeader(COLUMN_CR + 5) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + 5) + "" + ROW_SC_CR + ";" + getColumnHeader(COLUMN_CR + 1) + "" + (ROW_SC_CR + 1) + ":" + getColumnHeader(COLUMN_CR + 1) + "" + (ROW_SC_CR + 2) + ")");
}

function showFormAddNewCashIn() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCashIn')
    .setWidth(300)
    .setHeight(330);
  UI.showModalDialog(htmlOutput, 'New Cash In:');
}

function addNewCashIn(amount, date) {
  if (!ROW_EUROS_HTI || !COLLUMN_HTI) {
    console.log("|addNewCashIn| !ROW_EUROS_HTI || !COLLUMN_HTI");
    UI.alert('Erreur', "|addNewCashIn| !ROW_EUROS_HTI || !COLLUMN_HTI", UI.ButtonSet.OK);
    return;
  }

  SHEET.getRange(ROW_EUROS_HTI, COLLUMN_HTI, 3, 2).moveTo(SHEET.getRange((ROW_EUROS_HTI + 1), COLLUMN_HTI));

  let cell = SHEET.getRange(ROW_EUROS_HTI, COLLUMN_HTI);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setBackground(BG_GRIS_CLAIR_3);
  let dateFormat = "dd/MM/YYYY";
  cell.setNumberFormat(dateFormat);
  cell.setFontWeight("bold");
  cell.setFontSize(8);
  cell.setHorizontalAlignment("right");
  let dateComponents = date.split("-");
  let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
  cell.setValue(formattedDate);

  cell = SHEET.getRange(ROW_EUROS_HTI, (COLLUMN_HTI + 1))
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setFontSize(8);
  cell.setNumberFormat(getCellFormatNumberDollars(amount));
  cell.setValue(amount);

  addNewTransaction(amount, "Euros", amount, date, true);
}

function transaction(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedOptionSell, selectedQuantitySell, selectedDate) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let ui = SpreadsheetApp.getUi();

  let selectedQuantityBuyCalcul = getFormatCalculationScript(selectedQuantityBuy);
  let selectedAmountCalcul = getFormatCalculationScript(selectedAmount);
  let selectedQuantitySellCalcul = getFormatCalculationScript(selectedQuantitySell);

  addNewTransaction(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedDate, true);
  addNewTransaction(selectedAmount, selectedOptionSell, selectedQuantitySell, selectedDate, false);

  let chronologicalTransactionHistoric = findCellWithValueAllSheet("Historique Chronologique Transaction:");

  if (chronologicalTransactionHistoric) {
    let columnChronologicalTransactionHistoric = chronologicalTransactionHistoric.column;
    let lastRow = sheet.getLastRow();
    let rowCrypto = getRowIndexInColumnWithValue("Montant:", columnChronologicalTransactionHistoric);
    let cell;
    let rowLastTransaction = 0;
    if (rowCrypto) {
      console.log("rowCrypto: " + rowCrypto)
      console.log("columnChronologicalTransactionHistoric: " + columnChronologicalTransactionHistoric)
      let rangeSearch = sheet.getRange("" + getColumnHeader(columnChronologicalTransactionHistoric) + "" + rowCrypto + ":" + getColumnHeader(columnChronologicalTransactionHistoric) + "" + lastRow);
      let values = rangeSearch.getValues();
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === "") {
          break;
        }
        rowLastTransaction++;
      }
      rowLastTransaction = rowLastTransaction + rowCrypto;
      console.log("rowLastTransaction!!!!!!!! " + rowLastTransaction)
      cell = sheet.getRange(rowLastTransaction, columnChronologicalTransactionHistoric, 1, 8);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("right");
      cell.setFontSize(8);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 1), 1, 3);
      cell.setBackground(BG_VERT_CLAIR_3);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 4), 1, 3);
      cell.setBackground("#f4cccc");

      cell = sheet.getRange(rowLastTransaction, columnChronologicalTransactionHistoric);
      cell.setBackground(BG_BLEU_CLAIR_3)
      cell.setNumberFormat(getCellFormatNumberDollars(selectedAmountCalcul));
      cell.setValue(getFormatCell(selectedAmount));

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 1));
      cell.setValue(selectedOptionBuy);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 2));
      cell.setNumberFormat(getCellFormatNumber(selectedQuantityBuyCalcul));
      cell.setValue(getFormatCell(selectedQuantityBuy));

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 3));
      let average = selectedAmountCalcul / selectedQuantityBuyCalcul;
      console.log("Average: " + average)
      cell.setNumberFormat(getCellFormatNumberDollars(average));
      cell.setValue(average);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 4));
      cell.setValue(selectedOptionSell);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 5));
      cell.setNumberFormat(getCellFormatNumber(selectedQuantitySellCalcul));
      cell.setValue(getFormatCell(selectedQuantitySell));

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 6));
      average = selectedAmountCalcul / selectedQuantitySellCalcul;
      cell.setNumberFormat(getCellFormatNumberDollars(average));
      cell.setValue(average);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 7));
      let dateFormat = "dd/MM/YYYY";
      cell.setBackground(BG_BLEU_CLAIR_3)
      cell.setNumberFormat(dateFormat);
      cell.setValue(getCellDateFormat(selectedDate));


    } else {
      console.log("La valeur \"Crypto:\" n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur \"Crypto:\" n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur \"Historique Chronologique Transaction:\" n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur \"Historique Chronologique Transaction:\" n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function showRemoveCrypto() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formRemoveCrypto')
    .setWidth(300)
    .setHeight(250);
  htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
  UI.showModalDialog(htmlOutput, 'Supprimer une crypto:');
}

function removeCrypto(selectedOption) {
  if (!VALIDED_LOSS || !COLUMN_CR) {
    console.log("|removeCrypto| !VALIDED_LOSS || !COLUMN_CR");
    UI.alert('Erreur', "|removeCrypto| !VALIDED_LOSS || !COLUMN_CR", UI.ButtonSet.OK);
    return;
  }
  const activeRange = SHEET.getActiveRange();
  console.log("activeRange: " + activeRange.getA1Notation())
  const selectedOptionIndex = findCellsWithValueAndBackgroundAllSheet(selectedOption, "#d69bff");
  const selectedOptionRow_CR = getRowIndexInColumnWithValue(selectedOption, COLUMN_CR);
  let range, cell, value, values, validedLossCryptoRow, validedLossTotalRow, background;
  if (selectedOptionIndex && selectedOptionRow_CR) {
    cell = SHEET.getRange(selectedOptionRow_CR, (COLUMN_CR + 2))
    value = cell.getValue();

    if (value != 0) {
      let reponse = UI.alert(
        'Quantité Non Nulle',
        'Etes vous sûr de vouloir supprimer la cypto?',
        UI.ButtonSet.YES_NO);
      if (reponse == UI.Button.NO) {
        return;
      }
    }
    values = SHEET.getRange((VALIDED_LOSS.row + 3), VALIDED_LOSS.column, SHEET.getLastRow(), 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] == selectedOption) {
        validedLossCryptoRow = VALIDED_LOSS.row + 3 + i;
        break;
      }
      if (values[i][0] == "Total:") {
        validedLossTotalRow = VALIDED_LOSS.row + 3 + i;
        break;
      }
    }
    if (validedLossCryptoRow) {
      cell = SHEET.getRange(validedLossCryptoRow, VALIDED_LOSS.column + 1);
      cell.setValue(cell.getValue() + (-(SHEET.getRange(selectedOptionRow_CR, COLUMN_CR + 1).getValue())))
    }
    else if (validedLossTotalRow) {
      SHEET.getRange(validedLossTotalRow, VALIDED_LOSS.column, 1, 2).moveTo(SHEET.getRange((validedLossTotalRow + 1), VALIDED_LOSS.column));
      cell = SHEET.getRange(validedLossTotalRow, VALIDED_LOSS.column);
      cell.setBackground(BG_BLEU_CLAIR_3);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("left");
      cell.setFontSize(8);
      cell.setFontColor("#4285f4");
      cell.setValue(selectedOption);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      cell = SHEET.getRange(validedLossTotalRow, VALIDED_LOSS.column + 1);

      value = -(SHEET.getRange(selectedOptionRow_CR, COLUMN_CR + 1).getValue()); // Le "-" est important
      if (value < 0) {
        background = BG_ROUGE_CLAIR_3;
      } else {
        background = BG_VERT_CLAIR_3;
      }
      cell.setBackground(background);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("right");
      cell.setFontSize(8);
      cell.setNumberFormat(getCellFormatNumberDollars(value));
      cell.setValue(value);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      cell = SHEET.getRange((validedLossTotalRow + 1), VALIDED_LOSS.column + 1);
      cell.setFormula("=SUM(" + SHEET.getRange((VALIDED_LOSS.row + 3), VALIDED_LOSS.column + 1, (validedLossTotalRow - (VALIDED_LOSS.row + 3) + 1), 1).getA1Notation() + ")");


    }
    else {
      console.log("|removeCrypto| selectedOption et Total non trouvé");
      UI.alert('Erreur', "|removeCrypto| selectedOption et Total non trouvé", ui.ButtonSet.OK);
      return;
    }

    SHEET.getRange(1, selectedOptionIndex.column, SHEET.getLastRow(), 5).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    range = SHEET.getRange((selectedOptionRow_CR + 1), COLUMN_CR, (ROW_REVIEW_CR - (selectedOptionRow_CR + 1) + 1), 9)
    range.moveTo(SHEET.getRange((selectedOptionRow_CR), COLUMN_CR));

    const scRow_CR = getRowIndexInColumnWithValue("Stablecoin", COLUMN_CR);
    let plageFormatConditionnelle = SHEET.getRange("" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + 8) + "" + (scRow_CR - 1));
    let regleMiseEnForme = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + "<=0")
      .setBackground(BG_VERT_CLAIR_3)
      .setRanges([plageFormatConditionnelle])
      .build();
    SHEET.setConditionalFormatRules([regleMiseEnForme]);

    SHEET.setActiveSelection(activeRange);
  }
  else {
    console.log("|removeCrypto| selectedOptionIndex && selectedOptionIndex_CR");
    UI.alert('Erreur', "|removeCrypto| selectedOptionIndex && selectedOptionIndex_CR", ui.ButtonSet.OK);
  }
}
