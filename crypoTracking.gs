/*
-Implémenter etherscan tracker
*/


// Variables Feuilles -----------------------------------------------------------------------
const SHEET_NAME_TH = "HistoriqueTransaction";
const SHEET_NAME_MAIN = "Bilan";
const SHEET_NAME_DIAGRAM = "Diagramme";
const SHEET_NAME_OLD_CRYPTO = "AncienneCrypto";
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_MAIN);
const SHEET_HISTORIC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DIAGRAM);
const SHEET_TRANSACTION_HISTORIC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TH);
const SHEET_TRANSACTION_OLD_CRYPTO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_OLD_CRYPTO);
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const LAST_ROW = SHEET.getLastRow();
// -------------------------------------------------------------------------------------------

//---------- AUTRES Global --------------------------------------------------------------------
const BG_VERT_CLAIR_3 = "#d9ead3";
const BG_ROUGE_CLAIR_3 = "#f4cccc";
const BG_BLEU_CLAIR_3 = "#cfe2f3";
const BG_GRIS_CLAIR_3 = "#f3f3f3";
const STANDARD_FONT_SIZE = 8;
const FONT_COLOR_HEADER = "#351c75";
const ID_ETH_CMC = 1027
//--------------------------------------------------------------------------------------------

//--------- Value ----------------------------------------------------------------------------

// Crypto Review 
const CR = {
  VALUE_HEADER: "Bilan Crypto:",
  VALUE_BEFORE_FIRST_CRYPTO: "Crypto:",
  VALUE_AFTER_LAST_CRYPTO: "Total Investi:",
  VALUE_ROW_SC: "Stablecoin",
  VALUE_ROW_REVIEW: "Bilan:",
  VALUE_ROW_TOTAL_ASSETS: "Total Actif:",
  NB_FIELDS: 9,
  NB_COL_NAME: 0, // Nb colonne de différence avec la colonne de "CR.VALUE_HEADER"
  NB_COL_REMAIN_BET: 1,
  NB_COL_REMAIN_QUANTITY: 2,
  NB_COL_MEAN_BUY_PRICE: 3,
  NB_COL_CURRENT_PRICE: 4,
  NB_COL_VALUE: 5,
  NB_COL_REVIEW: 6,
  NB_COL_PNL: 7,
  NB_COL_DISTRIBUTION: 8,
  BG_CRYPTO_NAME: "#4285f4",
  FONT_COLOR_CRYPTO: "#4285f4",
}

// Transaction Historic
const TH = {
  NB_COL_AMOUNT: 0,
  NB_COL_QUANTITY: 1,
  NB_COL_MEAN_PRICE: 2,
  NB_COL_DATE: 3,
  NB_ROW_FIRST_TRANSACTION: 3,
  NB_COLUMN: 4,
  VALUE_SC: "Stablecoin",
  NB_FIELDS: 4,
  HEADER_FONT_SIZE: 20,
  BG_COLOR_FIELDS: "#ebcfff",
  FONT_COLOR_FIELDS: "#9900ff",
  FIRST_FIELD_NAME: "Montant:",
  SECOND_FIELD_NAME: "Quantité:",
  THIRD_FIELD_NAME: "Prix:",
  FOURTH_FIELD_NAME: "Date:",
  BG_CRYPTO_NAME: "#d69bff",
}

// NFT Review
const NR = {
  VALUE_HEADER: "Bilan NFT:",
  VALUE_ROW_TOTAL_INVESTED: "Total Investi:",
  NB_ROW_TO_MOVE: 2, 
  NB_FIELDS: 5,
  NB_COL_NAME: 0, 
  NB_COL_BUY_PRICE: 1,
  NB_COL_CURRENT_PRICE: 2,
  NB_COL_DATE: 3,
  NB_ROW_FIRST_NFT: 3,
  VALUE_BEFORE_FIRST_NFT: "NFT:",
}

// NFT Historic
const NH = {
  VALUE_HEADER: "Historique NFT Vendu:",
  VALUE_ROW_TOTAL: "Total:",
  NB_ROW_FIRST_NFT: 3,
  NB_FIELDS: 4,
  NB_COL_NAME: 0, 
  NB_COL_BUY_PRICE: 1,
  NB_COL_CURRENT_PRICE: 2,
  NB_COL_DATE: 3,
}

// History Total invested
const HTI = {
  VALUE_HEADER:"Historique Total Investi:",
  VALUE_ROW_TOTALD: "Total $:", // TOTALD pour DOLLARS
  VALUE_ROW_EUROS: "Euros to Usd Actuel:"
}

// Deleted Cryptos History
const DCH = {
  VALUE_HEADER: "Historique Cryptos Supprimées:",
  VALUE_LAST_ROW: "Total:",
}

// Chronological History Transaction
const CHT = {
  VALUE_HEADER: "Historique Chronologique Transaction:",
  VALUE_BEFORE_FIRST_TRANSAC: "Montant:",
  NB_ROW_FIRST_TRANSACTION: 3,
  NB_FIELDS: 8,
}

//--------------------------------------------------------------------------------------------
function refreshCryptoNftPrice(){
  callCmcApi()
  callOpenseaApi()
}
function geRowColGlobVar(sheetName, valueSearch, varGlobName) {
  let ui = SpreadsheetApp.getUi();
  let variable = SCRIPT_PROPERTIES.getProperty(varGlobName);
  if (variable) {
    let values = {};
    variable.split(";").forEach(keyValuePair => {
      let [key, value] = keyValuePair.split(":");
      values[key] = parseInt(value);
    });

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    if (sheet.getRange(values["row"], values["column"]).getValue() != valueSearch) {
      console.log("Mauvaise val row col")
      let indexValueSearch = findCellWithValueAllSheetName(valueSearch, sheetName)
      if (indexValueSearch) {
        let varGlobValue = "row:" + indexValueSearch.row + ";column:" + indexValueSearch.column
        SCRIPT_PROPERTIES.setProperty(varGlobName, varGlobValue)
        return indexValueSearch
      }
      else {
        console.log("|geRowColGlobVar| indexValueSearch undefined")
        ui.alert('Erreur', "|geRowColGlobVar| indexValueSearch undefined", ui.ButtonSet.OK)
        return
      }

    } else {
      console.log("Bonne val row col")
      return {
        row: values["row"],
        column: values["column"]
      }
    }

  }
  else {
    console.log("|geRowColGlobVar| La variable n'est pas présente dans les propriétés")
    ui.alert('Erreur', "|geRowColGlobVar| La variable n'est pas présente dans les propriétés", ui.ButtonSet.OK)
    return
  }

}

function onOpen() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const menuItems = [
    { name: 'Add New Transaction', functionName: 'showFormAddNewTransaction' },
    { name: 'Add New Crypto', functionName: 'showFormAddNewCrypto' },
    { name: 'Remove Crypto', functionName: 'showRemoveCrypto' },
    { name: 'Refresh Crypto And NFT Price', functionName: 'refreshCryptoNftPrice' },
    { name: 'Add Cash In', functionName: 'showFormAddNewCashIn' },
    { name: 'Add New NFT Transaction', functionName: 'showFormAddNewTransactionNFT' },
    { name: 'Maj All Crypto Data', functionName: 'majDataAllCrypto' },
    { name: 'See Review', functionName: 'main' },
  ];
  spreadsheet.addMenu('Cryptos Function', menuItems);
}

function showFormAddNewTransaction() {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);

  if (CRYPTO_REVIEW) {
    let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, CRYPTO_REVIEW.column) + 1;
    console.log("ROW_FIRST_CRYPTO_CR: " + ROW_FIRST_CRYPTO_CR);
    let LIST_ALL_CRYPTO = getAllCrypto(CRYPTO_REVIEW.column, ROW_FIRST_CRYPTO_CR);

    if (LIST_ALL_CRYPTO.length > 0 && ROW_FIRST_CRYPTO_CR) {
      let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewTransaction')
        .setWidth(600)
        .setHeight(900);

      htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
      ui.showModalDialog(htmlOutput, 'Ajouter une transaction:');
    }
    else {
      console.log("|showFormAddNewTransaction| LIST_ALL_CRYPTO.length > 0 && ROW_FIRST_CRYPTO_CR");
      ui.alert('Erreur', "|showFormAddNewTransaction| LIST_ALL_CRYPTO.length > 0 && ROW_FIRST_CRYPTO_CR", ui.ButtonSet.OK);
      return;
    }
  }
  else {
    console.log("|showFormAddNewTransaction| CRYPTO_REVIEW");
    ui.alert('Erreur', "|showFormAddNewTransaction| CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }
}

function transactionCrypto(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedOptionSell, selectedQuantitySell, selectedDate) {
  addNewTransaction(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedDate, true);
  addNewTransaction(selectedAmount, selectedOptionSell, selectedQuantitySell, selectedDate, false);
  addHistoricalTransactionHistory(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedOptionSell, selectedQuantitySell, selectedDate)
}

function addHistoricalTransactionHistory(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedOptionSell, selectedQuantitySell, selectedDate){
  let selectedQuantityBuyCalcul = getFormatCalculationScript(selectedQuantityBuy);
  let selectedAmountCalcul = getFormatCalculationScript(selectedAmount);
  let selectedQuantitySellCalcul = getFormatCalculationScript(selectedQuantitySell);

  let chronologicalTransactionHistoric = findCellWithValueAllSheetName(CHT.VALUE_HEADER, SHEET_NAME_TH);

  if (chronologicalTransactionHistoric) {
    let columnHCT = chronologicalTransactionHistoric.column;
    let rowFirstTransaction = getRowIndexInColumnWithValueWithSheet(SHEET_NAME_TH, CHT.VALUE_BEFORE_FIRST_TRANSAC, columnHCT) + 1
    let cell;
    let rowLastTransaction = rowFirstTransaction - 1 // Car on veut compter la première transaction qu'une fois
    if (rowFirstTransaction) {
       
      let rangeSearch = SHEET_TRANSACTION_HISTORIC.getRange(rowFirstTransaction, columnHCT, LAST_ROW, 1)

      let values = rangeSearch.getValues();
      for (let i = 0; i < values.length; i++) {
        if (values[i][0] === "") {
          break;
        }
        rowLastTransaction++;
      }
      console.log("rowLastTransaction: "+rowLastTransaction)
      let nbTransactions = rowLastTransaction - rowFirstTransaction + 1

      let rangeTransactions = SHEET_TRANSACTION_HISTORIC.getRange(rowFirstTransaction, columnHCT, nbTransactions, CHT.NB_FIELDS)

      rangeTransactions.moveTo(SHEET_TRANSACTION_HISTORIC.getRange(rowFirstTransaction+1, columnHCT))

      let range = SHEET_TRANSACTION_HISTORIC.getRange(rowFirstTransaction, columnHCT, 1, CHT.NB_FIELDS);

      range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
      range.setFontWeight("bold");
      range.setHorizontalAlignment("right");
      range.setFontSize(STANDARD_FONT_SIZE);

      cell = range.getCell(1, 1);
      cell.setBackground(BG_BLEU_CLAIR_3)
      cell.setNumberFormat(getCellFormatNumberDollars(selectedAmountCalcul));
      cell.setValue(getFormatCell(selectedAmount));

      cell = range.getCell(1, 2);
      cell.setValue(selectedOptionBuy);
      cell.setBackground(BG_VERT_CLAIR_3);

      cell = range.getCell(1, 3);
      cell.setNumberFormat(getCellFormatNumber(selectedQuantityBuyCalcul));
      cell.setValue(getFormatCell(selectedQuantityBuy));
      cell.setBackground(BG_VERT_CLAIR_3);

      cell = range.getCell(1, 4);
      let average = selectedAmountCalcul / selectedQuantityBuyCalcul;
      cell.setNumberFormat(getCellFormatNumberDollars(average));
      cell.setValue(average);
      cell.setBackground(BG_VERT_CLAIR_3);

      cell = range.getCell(1, 5);
      cell.setValue(selectedOptionSell);
      cell.setBackground(BG_ROUGE_CLAIR_3);

      cell = range.getCell(1, 6);
      cell.setNumberFormat(getCellFormatNumber(selectedQuantitySellCalcul));
      cell.setValue(getFormatCell(selectedQuantitySell));
      cell.setBackground(BG_ROUGE_CLAIR_3);

      cell = range.getCell(1, 7);
      average = selectedAmountCalcul / selectedQuantitySellCalcul;
      cell.setNumberFormat(getCellFormatNumberDollars(average));
      cell.setValue(average);
      cell.setBackground(BG_ROUGE_CLAIR_3);

      cell = range.getCell(1, 8);
      let dateFormat = "dd/MM/YYYY";
      cell.setBackground(BG_BLEU_CLAIR_3)
      cell.setNumberFormat(dateFormat);
      cell.setValue(getCellDateFormat(selectedDate));


    } else {
      console.log("|transaction| La valeur n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "|transaction| La valeur n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }
  } else {
    console.log("|transaction| La valeur n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "|transaction| La valeur n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function addNewTransaction(selectedAmount, selectedOption, selectedQuantity, selectedDate, isBuy) {
  let ui = SpreadsheetApp.getUi()
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  if (!CRYPTO_REVIEW) {
    console.log("|addNewTransaction| !CRYPTO_REVIEW")
    ui.alert('Erreur', "|addNewTransaction| !CRYPTO_REVIEW", ui.ButtonSet.OK)
    return
  }
  let COLUMN_CR = CRYPTO_REVIEW.column
  let background = BG_VERT_CLAIR_3
  let cell

  let selectedAmountCell = selectedAmount;
  if (selectedAmount.indexOf('.') !== -1) {
    selectedAmountCell = selectedAmount.replace('.', ',')
  } else if (selectedAmount.indexOf(',') !== -1) {
    selectedAmount = selectedAmount.replace(',', '.')
  }

  let selectedQuantityCell = selectedQuantity
  if (selectedQuantity.indexOf('.') !== -1) {
    selectedQuantityCell = selectedQuantity.replace('.', ',')
  } else if (selectedQuantity.indexOf(',') !== -1) {
    selectedQuantity = selectedQuantity.replace(',', '.')
  }

  let selectedOptionIndex = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, selectedOption, TH.BG_CRYPTO_NAME)
  let rowNewTransaction = selectedOptionIndex.row + TH.NB_ROW_FIRST_TRANSACTION
  if (!selectedOptionIndex) {
    console.log("|addNewTransaction| " + selectedOption + "n'a pas été trouvé")
    ui.alert('Erreur', "|addNewTransaction| " + selectedOption + "n'a pas été trouvé", ui.ButtonSet.OK)
    return;
  }

  if (isBuy === false) {
    background = BG_ROUGE_CLAIR_3
  }

  let values = SHEET_TRANSACTION_HISTORIC.getRange(rowNewTransaction, selectedOptionIndex.column, LAST_ROW, 1).getValues()
  let nbTransactions
  for (let j = 0; j < values.length; j++) {
    if (values[j][0] === "") {
      nbTransactions = j + TH.NB_ROW_FIRST_TRANSACTION + selectedOptionIndex.row
      break
    }
  }
  let nbRowToMove = nbTransactions - (TH.NB_ROW_FIRST_TRANSACTION + selectedOptionIndex.row) + 1
  let rangeTransactions = SHEET_TRANSACTION_HISTORIC.getRange(rowNewTransaction, selectedOptionIndex.column, nbRowToMove, TH.NB_COLUMN)

  rangeTransactions.moveTo(SHEET_TRANSACTION_HISTORIC.getRange(rowNewTransaction + 1, rangeTransactions.getColumn()))

  let range = SHEET_TRANSACTION_HISTORIC.getRange(rowNewTransaction, selectedOptionIndex.column, 1, TH.NB_COLUMN)
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID)
  range.setBackground(background)
  range.setHorizontalAlignment("right")
  range.setFontWeight("bold")
  range.setFontSize(STANDARD_FONT_SIZE)

  cell = range.getCell(1, TH.NB_COL_AMOUNT + 1)
  cell.setNumberFormat(getCellFormatNumberDollars(selectedAmount))
  cell.setValue(selectedAmountCell)

  cell = range.getCell(1, TH.NB_COL_QUANTITY + 1)
  cell.setNumberFormat(getCellFormatNumber(selectedQuantity))
  cell.setValue(selectedQuantityCell)

  cell = range.getCell(1, TH.NB_COL_MEAN_PRICE + 1)
  let averageBuyingPrice = selectedAmount / selectedQuantity
  cell.setNumberFormat(getCellFormatNumberDollars(averageBuyingPrice))
  cell.setValue(averageBuyingPrice);

  cell = range.getCell(1, TH.NB_COL_DATE + 1)
  let dateFormat = "dd/MM/YYYY"
  cell.setNumberFormat(dateFormat)
  let dateComponents = selectedDate.split("-")
  let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0]
  cell.setValue(formattedDate)

  let sellSum = 0
  let buySum = 0
  let buyQuantity = 0
  let sellQuantity = 0

  range = SHEET_TRANSACTION_HISTORIC.getRange(rowNewTransaction, selectedOptionIndex.column, LAST_ROW, TH.NB_COLUMN)

  values = range.getValues()
  const backgrounds = range.getBackgrounds()

  for (let i = 0; i < values.length; i++) {
    if (backgrounds[i][0] == BG_VERT_CLAIR_3) {
      buySum += values[i][0]
      buyQuantity += values[i][1]
    } else if (backgrounds[i][0] == BG_ROUGE_CLAIR_3) {
      sellSum += values[i][0]
      sellQuantity += values[i][1]
    } else {
      break
    }
  }
  let selectedOptionCR_Row = getRowIndexInColumnWithValue(selectedOption, COLUMN_CR)

  if (!selectedOptionCR_Row) {
    console.log("|addNewTransaction| !selectedOptionCR_Row")
    ui.alert('Erreur', "|addNewTransaction| !selectedOptionCR_Row", ui.ButtonSet.OK)
    return
  }

  cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + CR.NB_COL_REMAIN_BET))
  cell.setNumberFormat(getCellFormatNumberDollars((buySum - sellSum)))
  cell.setValue((buySum - sellSum))

  if (!(selectedOption == "Stablecoin")) {
    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + CR.NB_COL_REMAIN_QUANTITY))
    cell.setNumberFormat(getCellFormatNumber((buyQuantity - sellQuantity)))
    cell.setValue((buyQuantity - sellQuantity))

    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + CR.NB_COL_MEAN_BUY_PRICE))
    cell.setNumberFormat(getCellFormatNumberDollars((buySum / buyQuantity)))
    cell.setValue((buySum / buyQuantity))

    cell = SHEET.getRange(selectedOptionCR_Row, (COLUMN_CR + CR.NB_COL_PNL))
    cell.setFormula("IF(" + sellSum.toFixed(0) + "+" + SHEET.getRange(selectedOptionCR_Row, COLUMN_CR + CR.NB_COL_VALUE).getA1Notation() + ">" + buySum.toFixed(0) + ";((" + sellSum.toFixed(0) + "+" + SHEET.getRange(selectedOptionCR_Row, COLUMN_CR + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")-1;-(1-((" + sellSum.toFixed(0) + "+" + SHEET.getRange(selectedOptionCR_Row, COLUMN_CR + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")))")
  }

  descendingSortCR()
}


function callCmcApi() {
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  //refreshEurosPrice();

  if (!CRYPTO_REVIEW) {
    console.log("|callCmcApi| !CRYPTO_REVIEW");
    return;
  }
  let COLUMN_CR = CRYPTO_REVIEW.column;
  let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, COLUMN_CR) + 1;
  if (!COLUMN_CR || !ROW_FIRST_CRYPTO_CR) {
    console.log("|callCmcApi| !COLUMN_CR || !ROW_FIRST_CRYPTO_CR");
    return;
  }

  let cell;
  let mapIdCryptoGlobalString = SCRIPT_PROPERTIES.getProperty('mapIdCryptoGlobal');

  // Convertir la chaîne en map
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    mapIdCryptoGlobal[key] = value;
  });

  const startRow = ROW_FIRST_CRYPTO_CR;
  const urlSymbols = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
  const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  const apiKeyList = getApiKeyCmcList();
  let indexTabApiKey = 0;
  const indexColumnCurrentPrice = COLUMN_CR + CR.NB_COL_CURRENT_PRICE;
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
  for (let row = startRow; row <= LAST_ROW; row++) {
    const cryptoName = SHEET.getRange(row, COLUMN_CR).getValue();
    isCallValid = false;
    if (cryptoName === CR.VALUE_ROW_SC) {
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
          if (cryptoPrice != 0) {
            cell = SHEET.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice);
            console.log("cryptoPrice: " + cryptoPrice)
            console.log("getCellFormatNumberDollars(cryptoPrice): " + getCellFormatNumberDollars(cryptoPrice))
            cell.setNumberFormat(getCellFormatNumberDollars(cryptoPrice));
            cell.setFontWeight("bold");
            cell.setHorizontalAlignment("right");
            cell.setBackground(BG_GRIS_CLAIR_3);
            cell.setValue(cryptoPrice);
          }
        }
      }
    }

  }
  else {
    console.log(`Erreur lors de la récupération des données de la cryptomonnaie '${cryptoName}' : ${quoteResponse.getResponseCode()} - ${quoteJsonData}`);
  }
  descendingSortCR();
}

function showFormAddNewTransactionNFT() {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  if (!CRYPTO_REVIEW) {
    console.log("|showFormAddNewTransactionNFT| !CRYPTO_REVIEW");
    ui.alert('Erreur', "|showFormAddNewTransactionNFT| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }

  let COLUMN_CR = CRYPTO_REVIEW.column;
  let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, COLUMN_CR) + 1;
  let LIST_ALL_CRYPTO = getAllCrypto(CRYPTO_REVIEW.column, ROW_FIRST_CRYPTO_CR);
  if (!ROW_FIRST_CRYPTO_CR) {
    console.log("|showFormAddNewTransactionNFT| !ROW_FIRST_CRYPTO_CR");
    ui.alert('Erreur', "|showFormAddNewTransactionNFT| !ROW_FIRST_CRYPTO_CR", ui.ButtonSet.OK);
    return;
  }

  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewTransactionNFT')
    .setWidth(600)
    .setHeight(900);

  htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
  ui.showModalDialog(htmlOutput, 'Nouvelle Transaction NFT:');
}

function addNewTransactionNFT(idNftInput, optionSelect, amountInput, quantityInput, selectedDate, transactionType) {
  let ui = SpreadsheetApp.getUi();
  let isBuy = true;
  let cell;
  const NFT_REVIEW = findCellWithValueAllSheet(NR.VALUE_HEADER);
  const NFT_HISTORIC = findCellWithValueAllSheet(NH.VALUE_HEADER);
  let COLUMN_NFT_REVIEW = NFT_REVIEW.column;
  let ROW_TOTAL_INVESTED_NR = getRowIndexInColumnWithValue(NR.VALUE_ROW_TOTAL_INVESTED, COLUMN_NFT_REVIEW);


  if (!NFT_REVIEW || !NFT_HISTORIC || !ROW_TOTAL_INVESTED_NR) {
    console.log("|addNewTransactionNFT| !NFT_REVIEW || !NFT_HISTORIC || !ROW_TOTAL_INVESTED_NR");
    ui.alert('Erreur', "|addNewTransactionNFT| !NFT_REVIEW || !NFT_HISTORIC || !ROW_TOTAL_INVESTED_NR", ui.ButtonSet.OK);
    return;
  }
  let rangeNewNFT = SHEET.getRange(NFT_REVIEW.row + 3, COLUMN_NFT_REVIEW, 1, NR.NB_FIELDS)
  let nbRowToMove = (ROW_TOTAL_INVESTED_NR + 1) - (NFT_REVIEW.row + 3) + 1
  let dateFormat = "dd/MM/YYYY";
  let dateComponents = selectedDate.split("-");
  let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];

  let amountInputCell = amountInput;
  if (amountInput.indexOf('.') !== -1) {
    amountInputCell = amountInput.replace('.', ',');
  } else if (amountInput.indexOf(',') !== -1) {
    amountInput = amountInput.replace(',', '.');
  }

  if (transactionType == "achat") {
    isBuy = false;

    SHEET.getRange(NFT_REVIEW.row + 3, COLUMN_NFT_REVIEW, nbRowToMove, NR.NB_FIELDS).moveTo(SHEET.getRange(NFT_REVIEW.row + 4, COLUMN_NFT_REVIEW));

    rangeNewNFT.setBackground(BG_GRIS_CLAIR_3);
    rangeNewNFT.setFontWeight("bold");
    rangeNewNFT.setFontSize(STANDARD_FONT_SIZE);
    rangeNewNFT.setHorizontalAlignment("right");
    rangeNewNFT.setFontSize(STANDARD_FONT_SIZE);
    rangeNewNFT.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    cell = rangeNewNFT.getCell(1, 1)
    cell.setBackground(BG_BLEU_CLAIR_3);
    cell.setFontWeight("bold");
    cell.setHorizontalAlignment("left");
    cell.setFontSize(STANDARD_FONT_SIZE);
    cell.setFontColor("#4285f4");
    cell.setValue(idNftInput);

    cell = rangeNewNFT.getCell(1, 2)
    cell.setBackground(BG_GRIS_CLAIR_3);
    cell.setNumberFormat(getCellFormatNumberDollars(amountInput));
    cell.setValue(amountInputCell);

    cell = rangeNewNFT.getCell(1, 3)
    cell.setBackground(BG_GRIS_CLAIR_3);
    cell.setNumberFormat(getCellFormatNumberDollars(amountInput)); // On suppose qu'au début le prix estimé est égale au prix d'achat
    cell.setValue(amountInputCell);

    cell = rangeNewNFT.getCell(1, 4)
    cell.setBackground(BG_GRIS_CLAIR_3);
    cell.setNumberFormat("0$");
    cell.setValue(0);


    cell = rangeNewNFT.getCell(1, 5)
    cell.setBackground(BG_GRIS_CLAIR_3);
    cell.setNumberFormat(dateFormat);
    cell.setValue(formattedDate);

    cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR + 1, COLUMN_NFT_REVIEW + 1)
    cell.setFormula("=SUM(" + SHEET.getRange((NFT_REVIEW.row + NR.NB_ROW_FIRST_NFT), COLUMN_NFT_REVIEW + NR.NB_COL_BUY_PRICE, (ROW_TOTAL_INVESTED_NR - (NFT_REVIEW.row + NR.NB_ROW_FIRST_NFT) + 1), 1).getA1Notation() + ")");

    cell = SHEET.getRange(ROW_TOTAL_INVESTED_NR + 2, COLUMN_NFT_REVIEW + 1)
    cell.setFormula("=SUM(" + SHEET.getRange((NFT_REVIEW.row + NR.NB_ROW_FIRST_NFT), COLUMN_NFT_REVIEW + NR.NB_COL_CURRENT_PRICE, (ROW_TOTAL_INVESTED_NR - (NFT_REVIEW.row + NR.NB_ROW_FIRST_NFT) + 1), 1).getA1Notation() + ")-" + SHEET.getRange(ROW_TOTAL_INVESTED_NR + 1, COLUMN_NFT_REVIEW + 1).getA1Notation());

  } else {
    let ROW_NFT_SELECTED = getRowIndexInColumnWithValue(idNftInput, COLUMN_NFT_REVIEW);
    if (!ROW_NFT_SELECTED) {
      console.log("|addNewTransactionNFT| !ROW_NFT_SELECTED");
      ui.alert('Erreur', "|addNewTransactionNFT| !ROW_NFT_SELECTED", ui.ButtonSet.OK);
      return;
    }
    if (ROW_NFT_SELECTED == ROW_TOTAL_INVESTED_NR - 1) {
      SHEET.getRange(ROW_TOTAL_INVESTED_NR, COLUMN_NFT_REVIEW + 1).setValue(0);
      SHEET.getRange(ROW_TOTAL_INVESTED_NR + 1, COLUMN_NFT_REVIEW + 1).setValue(0);
    }
    let purchaseBuy = SHEET.getRange(ROW_NFT_SELECTED, COLUMN_NFT_REVIEW + NR.NB_COL_BUY_PRICE).getValue();
    SHEET.getRange(ROW_NFT_SELECTED + 1, COLUMN_NFT_REVIEW, (ROW_TOTAL_INVESTED_NR - ROW_NFT_SELECTED + 1), NR.NB_FIELDS).moveTo(SHEET.getRange(ROW_NFT_SELECTED, COLUMN_NFT_REVIEW));


    let rowHistoricNFT = NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT;
    let rowTotalHistoricNft = getRowIndexInColumnWithValue(NH.VALUE_ROW_TOTAL, NFT_HISTORIC.column)
    let nbRowToMoveNftHisto = rowTotalHistoricNft - (NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT) + 1
    let rangeHistoricNft = SHEET.getRange(NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT, NFT_HISTORIC.column, 1, NH.NB_FIELDS)

    SHEET.getRange(NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT, NFT_HISTORIC.column, nbRowToMoveNftHisto, NH.NB_FIELDS).moveTo(SHEET.getRange(NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT + 1, NFT_HISTORIC.column))
    let reviewNFT = (amountInput - purchaseBuy);

    cell = SHEET.getRange(rowHistoricNFT, NFT_HISTORIC.column, 1, NH.NB_FIELDS);
    cell.setBackground(BG_GRIS_CLAIR_3);



    rangeHistoricNft.setFontWeight("bold");
    rangeHistoricNft.setHorizontalAlignment("right");
    rangeHistoricNft.setFontSize(STANDARD_FONT_SIZE);
    rangeHistoricNft.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    cell = rangeHistoricNft.getCell(1, 1)
    cell.setHorizontalAlignment("left");
    cell.setBackground(BG_BLEU_CLAIR_3)
    cell.setFontColor(CR.FONT_COLOR_CRYPTO)
    cell.setValue(idNftInput);

    cell = rangeHistoricNft.getCell(1, 2)
    cell.setNumberFormat(getCellFormatNumberDollars(amountInput));
    cell.setValue(amountInputCell);

    cell = rangeHistoricNft.getCell(1, 3)
    cell.setNumberFormat('[Color50]+0$;[RED]-0$');
    cell.setValue(reviewNFT);

    cell = cell = rangeHistoricNft.getCell(1, 4)
    cell.setNumberFormat(dateFormat);
    cell.setValue(formattedDate);

    console.log("nbRowToMoveNftHisto: " + nbRowToMoveNftHisto)

    cell = SHEET.getRange(rowTotalHistoricNft + 1, NFT_HISTORIC.column + 1);
    cell.setNumberFormat('[Color50]+0$;[RED]-0$');
    cell.setFormula("=SUM(" + SHEET.getRange((NFT_HISTORIC.row + NH.NB_ROW_FIRST_NFT), NFT_HISTORIC.column + NH.NB_COL_CURRENT_PRICE, nbRowToMoveNftHisto, 1).getA1Notation() + ")");

  }

  addNewTransaction(amountInput, optionSelect, quantityInput, selectedDate, isBuy);
}

function showFormAddNewCrypto() {
  let ui = SpreadsheetApp.getUi();
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCrypto')
    .setWidth(300)
    .setHeight(220);

  ui.showModalDialog(htmlOutput, 'Ajouter une nouvelle crypto:');
}

function getAllInfosCryptoTransac(sheetName, crytoName) {
  let ui = SpreadsheetApp.getUi();
  let cryptoIndex = findCellsWithValueAndBackgroundAllSheet(sheetName, crytoName, TH.BG_CRYPTO_NAME);
  let currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!cryptoIndex) {
    console.log("|getAllInfosCryptoTransac| !cryptoIndex");
    ui.alert('Erreur', "|getAllInfosCryptoTransac| !cryptoIndex", ui.ButtonSet.OK);
    return;
  }

  let sellSum = 0;
  let buySum = 0;
  let buyQuantity = 0;
  let sellQuantity = 0;

  let rangeTransaction = currentSheet.getRange((cryptoIndex.row + 3), cryptoIndex.column, currentSheet.getLastRow(), 4);
  let valuesTransaction = rangeTransaction.getValues();
  let backgrounds = rangeTransaction.getBackgrounds();

  for (let i = 0; i < valuesTransaction.length; i++) {
    if (backgrounds[i][0] == BG_VERT_CLAIR_3) {
      buySum += valuesTransaction[i][0];
      buyQuantity += valuesTransaction[i][1];
    } else if (backgrounds[i][0] == BG_ROUGE_CLAIR_3) {
      sellSum += valuesTransaction[i][0];
      sellQuantity += valuesTransaction[i][1];
    } else {
      break;
    }
  }
  return {
    buySum: buySum,
    buyQuantity: buyQuantity,
    sellSum: sellSum,
    sellQuantity: sellQuantity
  }
}

function addNewCrypto(cryptoName) {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  if (!CRYPTO_REVIEW) {
    console.log("|addNewCrypto| !CRYPTO_REVIEW");
    ui.alert('Erreur', "|addNewCrypto| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }
  let COLUMN_CR = CRYPTO_REVIEW.column;
  let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, COLUMN_CR) + 1;
  let ROW_SC_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_SC, COLUMN_CR);
  let ROW_REVIEW_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_REVIEW, COLUMN_CR);
  let ROW_TOTAL_ASSETS_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_TOTAL_ASSETS, COLUMN_CR);

  if (!COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_SC_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR) {
    console.log("|addNewCrypto| !COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_SC_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR");
    ui.alert('Erreur', "|addNewCrypto| !COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_SC_CR || !ROW_REVIEW_CR || !ROW_TOTAL_ASSETS_CR", ui.ButtonSet.OK);
    return;
  }

  let range, cell;
  const stablecoinTH = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, TH.VALUE_SC, TH.BG_CRYPTO_NAME);
  let remainingBet = 0;
  let remainingQuantity = 0
  SHEET_TRANSACTION_HISTORIC.insertColumnsBefore((stablecoinTH.column - 1), TH.NB_FIELDS + 1); // -1 -> car sinon ca prend le format de "StableCoin Bincance" (cellule colorées etc))

  let sourceRange = getRangeCrypto(SHEET_NAME_OLD_CRYPTO, cryptoName)
  let cryptoIndex = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_OLD_CRYPTO, cryptoName, TH.BG_CRYPTO_NAME)
  if (cryptoIndex) {
    // Copier la liste des transactions de la crypto depuis "AncienneCrypto"
    targetRange = SHEET_TRANSACTION_HISTORIC.getRange(stablecoinTH.row, stablecoinTH.column, sourceRange.getNumRows(), sourceRange.getNumColumns());
    sourceRange.copyTo(targetRange, { contentsOnly: true });
    sourceRange.copyTo(targetRange, { formatOnly: true });
    //let cryptoInfos = getAllInfosCryptoTransac(SHEET_NAME_OLD_CRYPTO, cryptoName)
    //remainingBet = cryptoInfos.buySum - cryptoInfos.sellSum;
    //remainingQuantity = cryptoInfos.buyQuantity - cryptoInfos.sellQuantity
    // Supprimer la crypto de la liste des cryptos supp
    const VALIDED_LOSS = findCellWithValueAllSheet(DCH.VALUE_HEADER);
    let rowCryptoVL = getRowIndexInColumnWithValue(cryptoName, VALIDED_LOSS.column)
    let rowTotalVL = getRowIndexInColumnWithValue(DCH.VALUE_LAST_ROW, VALIDED_LOSS.column)
    SHEET.getRange(rowCryptoVL + 1, VALIDED_LOSS.column, rowTotalVL - rowCryptoVL, 2).moveTo(SHEET.getRange(rowCryptoVL, VALIDED_LOSS.column));

    // Supprimer la crypto de la feuille AncienneCrypto
    SHEET_TRANSACTION_OLD_CRYPTO.deleteColumns(sourceRange.getColumn(), sourceRange.getNumColumns() + 1)
  }
  else {
    // Header
    range = SHEET_TRANSACTION_HISTORIC.getRange(stablecoinTH.row, stablecoinTH.column, 2, TH.NB_FIELDS); // 2 --> Taille Header
    range.setBackground(TH.BG_CRYPTO_NAME);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setFontSize(TH.HEADER_FONT_SIZE);
    range.setFontColor(FONT_COLOR_HEADER);
    range.setValue(cryptoName);
    range.merge();
    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    range = SHEET_TRANSACTION_HISTORIC.getRange((stablecoinTH.row + 2), stablecoinTH.column, 1, TH.NB_FIELDS); // 2 --> diff entre le sommet du header et ligne paramètre (date...)

    range.setBackground(TH.BG_COLOR_FIELDS);
    range.setFontColor(TH.FONT_COLOR_FIELDS);
    range.setFontWeight("bold");
    range.setFontSize(STANDARD_FONT_SIZE);
    range.setHorizontalAlignment("left");
    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    cell = range.getCell(1, TH.NB_COL_AMOUNT + 1);
    cell.setValue(TH.FIRST_FIELD_NAME)

    cell = range.getCell(1, TH.NB_COL_QUANTITY + 1);
    cell.setValue(TH.SECOND_FIELD_NAME)

    cell = range.getCell(1, TH.NB_COL_MEAN_PRICE + 1);
    cell.setValue(TH.THIRD_FIELD_NAME)

    cell = range.getCell(1, TH.NB_COL_DATE + 1);
    cell.setValue(TH.FOURTH_FIELD_NAME);

  }

  SHEET.getRange(ROW_SC_CR, COLUMN_CR, (ROW_REVIEW_CR - ROW_SC_CR + 1), CR.NB_FIELDS).moveTo(SHEET.getRange((ROW_SC_CR + 1), COLUMN_CR));

  cell = SHEET.getRange(ROW_SC_CR, COLUMN_CR);
  cell.setBackground(BG_BLEU_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("left");
  cell.setFontSize(STANDARD_FONT_SIZE);


  cell.setFontColor(CR.FONT_COLOR_CRYPTO);
  cell.setValue(cryptoName);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + 1), 1, CR.NB_FIELDS - 1);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("right");
  cell.setFontSize(STANDARD_FONT_SIZE);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_REMAIN_BET));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_REMAIN_QUANTITY));
  cell.setValue(0);

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_MEAN_BUY_PRICE));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_CURRENT_PRICE));
  cell.setValue(0);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_VALUE));
  cell.setFormula("=" + getColumnHeader(COLUMN_CR + CR.NB_COL_MEAN_BUY_PRICE) + "" + ROW_SC_CR + "*" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_QUANTITY) + "" + ROW_SC_CR);
  cell.setNumberFormat('0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_REVIEW));

  cell.setFormula("=" + getColumnHeader(COLUMN_CR + CR.NB_COL_VALUE) + "" + ROW_SC_CR + "-" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_BET) + "" + ROW_SC_CR);
  cell.setNumberFormat('[Color50]+0$;[RED]-0$');

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_PNL));
  cell.setNumberFormat('[Color50]+0.00%;[Red]-0.00%');
  cell.setValue(0)

  cell = SHEET.getRange(ROW_SC_CR, (COLUMN_CR + CR.NB_COL_DISTRIBUTION));
  cell.setFormula("=" + getColumnHeader(COLUMN_CR + CR.NB_COL_VALUE) + "" + ROW_SC_CR + "/" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_BET) + "" + (ROW_TOTAL_ASSETS_CR + 1)); // +1 pour prendre en compte la ligne de la nouvelle crypto qui a été ajouté
  cell.setNumberFormat('0.00%');

  // Créer la règle de mise en forme conditionnelle
  let plageFormatConditionnelle = SHEET.getRange("" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_BET) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + CR.NB_COL_DISTRIBUTION) + "" + ROW_SC_CR);
  let regleMiseEnForme = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_BET) + "" + ROW_FIRST_CRYPTO_CR + "<=0")
    .setBackground(BG_VERT_CLAIR_3) // Couleur de fond verte en cas de condition satisfaite
    .setRanges([plageFormatConditionnelle])
    .build();
  SHEET.setConditionalFormatRules([regleMiseEnForme]);
  cell = SHEET.getRange((ROW_TOTAL_ASSETS_CR + 1), (COLUMN_CR + CR.NB_COL_REMAIN_BET)); // +1 --> prendre en compte l'ajout de la nouvelle crypto
  cell.setFormula("=SUM(" + getColumnHeader(COLUMN_CR + CR.NB_COL_VALUE) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + CR.NB_COL_VALUE) + "" + ROW_SC_CR + ";" + getColumnHeader(COLUMN_CR + 1) + "" + (ROW_SC_CR + 1) + ":" + getColumnHeader(COLUMN_CR + CR.NB_COL_REMAIN_BET) + "" + (ROW_SC_CR + 1) + ")");

  majDataCrypto(cryptoName)
  callCmcApi();
}

function showFormAddNewCashIn() {
  let ui = SpreadsheetApp.getUi();
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCashIn')
    .setWidth(300)
    .setHeight(330);
  ui.showModalDialog(htmlOutput, 'New Cash In:');
}

function addNewCashIn(amount, date) {
  let ui = SpreadsheetApp.getUi();
  const HISTORY_TOTAL_INVESTED = findCellWithValueAllSheet(HTI.VALUE_HEADER);
  if (!HISTORY_TOTAL_INVESTED) {
    console.log("|addNewCashIn| !HISTORY_TOTAL_INVESTED");
    ui.alert('Erreur', "|addNewCashIn| !HISTORY_TOTAL_INVESTED", ui.ButtonSet.OK);
    return;
  }

  let COLLUMN_HTI = HISTORY_TOTAL_INVESTED.column;
  let ROW_TOTALD_HTI = getRowIndexInColumnWithValue(HTI.VALUE_ROW_TOTALD, COLLUMN_HTI);
  let ammountInUSd = getEurosToUsdPrice() * amount;

  if (!ROW_TOTALD_HTI || !COLLUMN_HTI) {
    console.log("|addNewCashIn| !ROW_TOTALD_HTI || !COLLUMN_HTI");
    ui.alert('Erreur', "|addNewCashIn| !ROW_TOTALD_HTI || !COLLUMN_HTI", ui.ButtonSet.OK);
    return;
  }

  let nbRow = ROW_TOTALD_HTI - HISTORY_TOTAL_INVESTED.row - 2 // Pour ne pas prendre le header + nom colonne

  SHEET.getRange(HISTORY_TOTAL_INVESTED.row + 3, COLLUMN_HTI, nbRow, 3).moveTo(SHEET.getRange(HISTORY_TOTAL_INVESTED.row + 4, COLLUMN_HTI));

  let cell = SHEET.getRange(HISTORY_TOTAL_INVESTED.row + 3, COLLUMN_HTI);
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setBackground(BG_GRIS_CLAIR_3);
  let dateFormat = "dd/MM/YYYY";
  cell.setNumberFormat(dateFormat);
  cell.setFontWeight("bold");
  cell.setFontSize(STANDARD_FONT_SIZE);
  cell.setHorizontalAlignment("right");
  let dateComponents = date.split("-");
  let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
  cell.setValue(formattedDate);

  cell = SHEET.getRange(HISTORY_TOTAL_INVESTED.row + 3, (COLLUMN_HTI + 1))
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setFontSize(STANDARD_FONT_SIZE);
  cell.setNumberFormat("0€");
  cell.setValue(amount);

  cell = SHEET.getRange(HISTORY_TOTAL_INVESTED.row + 3, (COLLUMN_HTI + 2))
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setBackground(BG_GRIS_CLAIR_3);
  cell.setFontWeight("bold");
  cell.setFontSize(STANDARD_FONT_SIZE);
  cell.setNumberFormat("0.00$");
  cell.setValue(getEurosToUsdPrice());

  cell = SHEET.getRange(ROW_TOTALD_HTI, (COLLUMN_HTI + 1))
  cell.setFormula("=SUM(" + SHEET.getRange((HISTORY_TOTAL_INVESTED.row + 3), HISTORY_TOTAL_INVESTED.column + 1, ROW_TOTALD_HTI - 7, 1).getA1Notation() + ")");

  // Eur to Usd Moyen

  let amountValues = SHEET.getRange((HISTORY_TOTAL_INVESTED.row + 3), HISTORY_TOTAL_INVESTED.column + 1, (ROW_TOTALD_HTI - 3 - (HISTORY_TOTAL_INVESTED.row + 3) + 1), 1).getValues();
  let meanValues = SHEET.getRange((HISTORY_TOTAL_INVESTED.row + 3), HISTORY_TOTAL_INVESTED.column + 2, (ROW_TOTALD_HTI - 3 - (HISTORY_TOTAL_INVESTED.row + 3) + 1), 1).getValues();
  let amountEurSum = 0;
  let amountUsdSum = 0;
  for (let i = 0; i < amountValues.length; i++) {
    amountEurSum += amountValues[i][0];
    amountUsdSum += amountValues[i][0] * meanValues[i][0];
  }
  let meanEurToUsd = amountUsdSum / amountEurSum;

  cell = SHEET.getRange(ROW_TOTALD_HTI - 1, (COLLUMN_HTI + 1))
  cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  cell.setFontWeight("bold");
  cell.setFontSize(STANDARD_FONT_SIZE);
  cell.setNumberFormat("0.00$");
  cell.setValue(meanEurToUsd);

  addNewTransaction(ammountInUSd.toString(), "Stablecoin", ammountInUSd.toString(), date, true);
}

function showRemoveCrypto() {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  if (!CRYPTO_REVIEW) {
    console.log("|showRemoveCrypto| !CRYPTO_REVIEW");
    ui.alert('Erreur', "|showRemoveCrypto| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }
  let COLUMN_CR = CRYPTO_REVIEW.column;
  let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, COLUMN_CR) + 1;
  let LIST_ALL_CRYPTO = getAllCrypto(CRYPTO_REVIEW.column, ROW_FIRST_CRYPTO_CR);
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formRemoveCrypto')
    .setWidth(300)
    .setHeight(250);
  htmlOutput.append('<script>let choices = ' + JSON.stringify(LIST_ALL_CRYPTO) + ';</script>');
  ui.showModalDialog(htmlOutput, 'Supprimer une crypto:');
}

function removeCrypto(selectedOption) {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);
  const VALIDED_LOSS = findCellWithValueAllSheet(DCH.VALUE_HEADER);
  if (!CRYPTO_REVIEW || !VALIDED_LOSS) {
    console.log("|removeCrypto| !CRYPTO_REVIEW || !VALIDED_LOSS");
    ui.alert('Erreur', "|removeCrypto| !CRYPTO_REVIEW || !VALIDED_LOSS", ui.ButtonSet.OK);
    return;
  }
  let COLUMN_CR = CRYPTO_REVIEW.column;
  let ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, COLUMN_CR) + 1;
  let ROW_REVIEW_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_REVIEW, COLUMN_CR);

  if (!COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_REVIEW_CR) {
    console.log("|removeCrypto| !COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_REVIEW_CR");
    ui.alert('Erreur', "|removeCrypto| !COLUMN_CR || !ROW_FIRST_CRYPTO_CR || !ROW_REVIEW_CR", ui.ButtonSet.OK);
    return;
  }

  //const activeRange = SHEET.getActiveRange();
  const selectedOptionIndex = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, selectedOption, TH.BG_CRYPTO_NAME);
  const selectedOptionRow_CR = getRowIndexInColumnWithValue(selectedOption, COLUMN_CR);
  let range, cell, value, values, validedLossCryptoRow, validedLossTotalRow, background;
  if (selectedOptionIndex && selectedOptionRow_CR) {
    cell = SHEET.getRange(selectedOptionRow_CR, (COLUMN_CR + 2))
    value = cell.getValue();

    if (value != 0) {
      let reponse = ui.alert(
        'Quantité Non Nulle',
        'Etes vous sûr de vouloir supprimer la cypto?',
        ui.ButtonSet.YES_NO);
      if (reponse == ui.Button.NO) {
        return;
      }
    }
    values = SHEET.getRange((VALIDED_LOSS.row + 3), VALIDED_LOSS.column, LAST_ROW, 1).getValues();
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
      SHEET.getRange(VALIDED_LOSS.row + 3, VALIDED_LOSS.column, validedLossTotalRow - (VALIDED_LOSS.row + 3) + 1, 2).moveTo(SHEET.getRange(VALIDED_LOSS.row + 4, VALIDED_LOSS.column));
      cell = SHEET.getRange(VALIDED_LOSS.row + 3, VALIDED_LOSS.column);
      cell.setBackground(BG_BLEU_CLAIR_3);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("left");
      cell.setFontSize(STANDARD_FONT_SIZE);
      cell.setFontColor("#4285f4");
      cell.setValue(selectedOption);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      cell = SHEET.getRange(VALIDED_LOSS.row + 3, VALIDED_LOSS.column + 1);

      value = -(SHEET.getRange(selectedOptionRow_CR, COLUMN_CR + 1).getValue()); // Le "-" est important

      cell.setBackground(BG_GRIS_CLAIR_3);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("right");
      cell.setFontSize(STANDARD_FONT_SIZE);
      cell.setNumberFormat('[Color50]+0$;[RED]-0$');
      cell.setValue(value);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      cell = SHEET.getRange((validedLossTotalRow + 1), VALIDED_LOSS.column + 1);
      cell.setFormula("=SUM(" + SHEET.getRange((VALIDED_LOSS.row + 3), VALIDED_LOSS.column + 1, (validedLossTotalRow - (VALIDED_LOSS.row + 3) + 1), 1).getA1Notation() + ")");


    }
    else {
      console.log("|removeCrypto| selectedOption et Total non trouvé");
      ui.alert('Erreur', "|removeCrypto| selectedOption et Total non trouvé", ui.ButtonSet.OK);
      return;
    }

    addToOldCrypto(selectedOption);
    SHEET_TRANSACTION_HISTORIC.getRange(1, selectedOptionIndex.column, LAST_ROW, 5).deleteCells(SpreadsheetApp.Dimension.COLUMNS);


    range = SHEET.getRange((selectedOptionRow_CR + 1), COLUMN_CR, (ROW_REVIEW_CR - (selectedOptionRow_CR + 1) + 1), 9)
    range.moveTo(SHEET.getRange((selectedOptionRow_CR), COLUMN_CR));

    const scRow_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_SC, COLUMN_CR);
    let plageFormatConditionnelle = SHEET.getRange("" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + ":" + getColumnHeader(COLUMN_CR + 8) + "" + (scRow_CR - 1));
    let regleMiseEnForme = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$" + getColumnHeader(COLUMN_CR + 1) + "" + ROW_FIRST_CRYPTO_CR + "<=0")
      .setBackground(BG_VERT_CLAIR_3)
      .setRanges([plageFormatConditionnelle])
      .build();
    SHEET.setConditionalFormatRules([regleMiseEnForme]);

    //SHEET.setActiveSelection(activeRange);
  }
  else {
    console.log("|removeCrypto| selectedOptionIndex && selectedOptionIndex_CR");
    ui.alert('Erreur', "|removeCrypto| selectedOptionIndex && selectedOptionIndex_CR", ui.ButtonSet.OK);
  }
}


function descendingSortCR() {
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);

  if (!CRYPTO_REVIEW) {
    console.log("|descendingSortCR| !CRYPTO_REVIEW");
    return;
  }

  const ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, CRYPTO_REVIEW.column) + 1;
  const ROW_LAST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_SC, CRYPTO_REVIEW.column);
  if (!ROW_FIRST_CRYPTO_CR || !ROW_LAST_CRYPTO_CR) {
    console.log("|descendingSortCR| !ROW_FIRST_CRYPTO_CR || !ROW_LAST_CRYPTO_CR");
    return;
  }

  let cell;
  let allCryptoList = [];

  let nbRowRange = ROW_LAST_CRYPTO_CR - ROW_FIRST_CRYPTO_CR
  const range = SHEET.getRange(ROW_FIRST_CRYPTO_CR, CRYPTO_REVIEW.column, nbRowRange, CR.NB_FIELDS);
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    allCryptoList.push({
      name: values[i][CR.NB_COL_NAME],
      remainingBet: values[i][CR.NB_COL_REMAIN_BET],
      remainingQuantity: values[i][CR.NB_COL_REMAIN_QUANTITY],
      averagePurchasePrice: values[i][CR.NB_COL_MEAN_BUY_PRICE],
      currentPrice: values[i][CR.NB_COL_CURRENT_PRICE],
      value: values[i][CR.NB_COL_VALUE],
      review: values[i][CR.NB_COL_REVIEW],
      buySum: extractSumBuySellFormula(ROW_FIRST_CRYPTO_CR + i, CRYPTO_REVIEW.column + CR.NB_COL_PNL).buySum,
      sellSum: extractSumBuySellFormula(ROW_FIRST_CRYPTO_CR + i, CRYPTO_REVIEW.column + CR.NB_COL_PNL).sellSum,
      distribution: values[i][CR.NB_COL_DISTRIBUTION],
    });
  }


  for (let i = 1; i <= values.length; i++) {
    allCryptoList.sort(descendingComparisonFunction);
    cell = range.getCell(i, CR.NB_COL_NAME + 1);
    cell.setValue(allCryptoList[i - 1].name);

    cell = range.getCell(i, CR.NB_COL_REMAIN_BET + 1);
    cell.setNumberFormat(getCellFormatNumberDollars(allCryptoList[i - 1].remainingBet));
    cell.setValue(allCryptoList[i - 1].remainingBet);

    cell = range.getCell(i, CR.NB_COL_REMAIN_QUANTITY + 1);
    cell.setNumberFormat(getCellFormatNumber(allCryptoList[i - 1].remainingQuantity));
    cell.setValue(allCryptoList[i - 1].remainingQuantity);

    cell = range.getCell(i, CR.NB_COL_MEAN_BUY_PRICE + 1);
    cell.setNumberFormat(getCellFormatNumberDollars(allCryptoList[i - 1].averagePurchasePrice));
    cell.setValue(allCryptoList[i - 1].averagePurchasePrice);

    cell = range.getCell(i, CR.NB_COL_CURRENT_PRICE + 1);
    cell.setNumberFormat(getCellFormatNumberDollars(allCryptoList[i - 1].currentPrice));
    cell.setValue(allCryptoList[i - 1].currentPrice);

    cell = range.getCell(i, CR.NB_COL_VALUE + 1);
    cell.setNumberFormat(getCellFormatNumberDollars(allCryptoList[i - 1].value));
    cell.setFormula("=" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_CURRENT_PRICE) + "" + cell.getRow() + "*" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_REMAIN_QUANTITY) + "" + cell.getRow());
    cell.setNumberFormat('0$');

    cell = range.getCell(i, CR.NB_COL_REVIEW + 1);
    cell.setFormula("=" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_VALUE) + "" + cell.getRow() + "-" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_REMAIN_BET) + "" + cell.getRow());

    cell = range.getCell(i, CR.NB_COL_PNL + 1);

    if (allCryptoList[i - 1].buySum != 0) {
      cell.setFormula("IF(" + allCryptoList[i - 1].sellSum + "+" + SHEET.getRange(cell.getRow(), CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ">" + allCryptoList[i - 1].buySum + ";((" + allCryptoList[i - 1].sellSum + "+" + SHEET.getRange(cell.getRow(), CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + allCryptoList[i - 1].buySum + ")-1;-(1-((" + allCryptoList[i - 1].sellSum + "+" + SHEET.getRange(cell.getRow(), CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + allCryptoList[i - 1].buySum + ")))")
    }
    else {
      cell.setValue(0)
    }


    cell = range.getCell(i, CR.NB_COL_DISTRIBUTION + 1);
    cell.setFormula("=" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_VALUE) + "" + cell.getRow() + "/" + getColumnHeader(CRYPTO_REVIEW.column + CR.NB_COL_REMAIN_BET) + "" + (ROW_LAST_CRYPTO_CR + CR.NB_COL_REMAIN_QUANTITY));
  }
}

function majDataAllCrypto() {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);

  if (!CRYPTO_REVIEW) {
    console.log("|crDataMAJ| !CRYPTO_REVIEW");
    ui.alert('Erreur', "|crDataMAJ| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }

  const ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, CRYPTO_REVIEW.column) + 1;
  const ROW_LAST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_SC, CRYPTO_REVIEW.column);

  let cell;
  const rangeCR = SHEET.getRange(ROW_FIRST_CRYPTO_CR, CRYPTO_REVIEW.column, (ROW_LAST_CRYPTO_CR - ROW_FIRST_CRYPTO_CR), 9);
  const valuesCR = rangeCR.getValues();
  let valuesTransaction;
  let rangeTransaction;
  let indexCryptoTransaction;
  let sellSum;
  let buySum;
  let buyQuantity;
  let sellQuantity;
  let backgrounds;

  let remainingBet;
  let remainingQuantity;
  let averageBuyingPrice;
  let valueCrypto;
  let review;
  let pnl;
  let distribution;
  let currentPrice;

  const totalInvestedCR_Row = getRowIndexInColumnWithValue(CR.VALUE_ROW_TOTAL_ASSETS, CRYPTO_REVIEW.column);
  const valueTotalInvested = SHEET.getRange(totalInvestedCR_Row, CRYPTO_REVIEW.column + 1).getValue();

  for (let i = 0; i < valuesCR.length; i++) {
    indexCryptoTransaction = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, valuesCR[i][0], TH.BG_CRYPTO_NAME);
    if (!indexCryptoTransaction) {
      console.log("|crDataMAJ| !indexCryptoTransaction | Crypto: " + valuesCR[i][0]);
      ui.alert('Erreur', "|crDataMAJ| !indexCryptoTransaction | Crypto: " + valuesCR[i][0], ui.ButtonSet.OK);
    }

    sellSum = 0;
    buySum = 0;
    buyQuantity = 0;
    sellQuantity = 0;

    rangeTransaction = SHEET_TRANSACTION_HISTORIC.getRange((indexCryptoTransaction.row + 3), indexCryptoTransaction.column, SHEET_TRANSACTION_HISTORIC.getLastRow(), 4);
    valuesTransaction = rangeTransaction.getValues();
    backgrounds = rangeTransaction.getBackgrounds();

    for (let i = 0; i < valuesTransaction.length; i++) {
      if (backgrounds[i][0] == BG_VERT_CLAIR_3) {
        buySum += valuesTransaction[i][0];
        buyQuantity += valuesTransaction[i][1];
      } else if (backgrounds[i][0] == BG_ROUGE_CLAIR_3) {
        sellSum += valuesTransaction[i][0];
        sellQuantity += valuesTransaction[i][1];
      } else {
        break;
      }
    }

    currentPrice = rangeCR.getCell(i + 1, 5).getValue();
    remainingBet = buySum - sellSum;
    remainingQuantity = buyQuantity - sellQuantity;
    if (buyQuantity != 0) {
      averageBuyingPrice = buySum / buyQuantity
      if (currentPrice < averageBuyingPrice) {
        pnl = -(1 - (currentPrice / averageBuyingPrice));
      } else {
        pnl = (currentPrice / averageBuyingPrice) - 1;
      }
    } else {
      averageBuyingPrice = 0
      pnl = 0
    }

    valueCrypto = remainingQuantity * currentPrice;
    review = valueCrypto - remainingBet;

    distribution = valueCrypto / valueTotalInvested;

    console.log("crypto: " + valuesCR[i][0] + " | remainingBet: " + remainingBet + " | remainingQuantity: " + remainingQuantity + " | averageBuyingPrice: " + averageBuyingPrice + " | valueCrypto: " + valueCrypto + " | review: " + review + " | distribution: " + distribution + " | pnl: " + pnl);

    cell = rangeCR.getCell(i + 1, 2);
    cell.setNumberFormat(getCellFormatNumberDollars(remainingBet));
    cell.setValue(remainingBet);

    cell = rangeCR.getCell(i + 1, 3);
    cell.setNumberFormat(getCellFormatNumber(remainingQuantity));
    cell.setValue(remainingQuantity);

    cell = rangeCR.getCell(i + 1, 4);
    cell.setNumberFormat(getCellFormatNumberDollars(averageBuyingPrice));
    cell.setValue(averageBuyingPrice);

    cell = rangeCR.getCell(i + 1, 8);
    cell.setFormula("IF(" + sellSum.toFixed(0) + "+" + SHEET.getRange(i + 5, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ">" + buySum.toFixed(0) + ";((" + sellSum.toFixed(0) + "+" + SHEET.getRange(i + 5, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")-1;-(1-((" + sellSum.toFixed(0) + "+" + SHEET.getRange(i + 5, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")))")
  }
  descendingSortCR();
}

function majDataCrypto(cryptoName) {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);

  if (!CRYPTO_REVIEW) {
    console.log("|crDataMAJ| !CRYPTO_REVIEW");
    ui.alert('Erreur', "|crDataMAJ| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }

  let cell;
  let valuesTransaction;
  let rangeTransaction;
  let indexCryptoTransaction;
  let sellSum;
  let buySum;
  let buyQuantity;
  let sellQuantity;
  let backgrounds;

  let remainingBet;
  let remainingQuantity;
  let averageBuyingPrice;
  let valueCrypto;
  let review;
  let pnl;
  let distribution;
  let currentPrice;

  const totalInvestedCR_Row = getRowIndexInColumnWithValue(CR.VALUE_ROW_TOTAL_ASSETS, CRYPTO_REVIEW.column);
  const valueTotalInvested = SHEET.getRange(totalInvestedCR_Row, CRYPTO_REVIEW.column + 1).getValue();
  const rowCryptoCR = getRowIndexInColumnWithValue(cryptoName, CRYPTO_REVIEW.column)

  indexCryptoTransaction = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, cryptoName, TH.BG_CRYPTO_NAME);
  if (!indexCryptoTransaction) {
    console.log("|crDataMAJ| !indexCryptoTransaction | Crypto: " + cryptoName);
    ui.alert('Erreur', "|crDataMAJ| !indexCryptoTransaction | Crypto: " + cryptoName, ui.ButtonSet.OK);
  }

  sellSum = 0;
  buySum = 0;
  buyQuantity = 0;
  sellQuantity = 0;

  rangeTransaction = SHEET_TRANSACTION_HISTORIC.getRange((indexCryptoTransaction.row + 3), indexCryptoTransaction.column, SHEET_TRANSACTION_HISTORIC.getLastRow(), 4);
  valuesTransaction = rangeTransaction.getValues();
  backgrounds = rangeTransaction.getBackgrounds();

  for (let i = 0; i < valuesTransaction.length; i++) {
    if (backgrounds[i][0] == BG_VERT_CLAIR_3) {
      buySum += valuesTransaction[i][0];
      buyQuantity += valuesTransaction[i][1];
    } else if (backgrounds[i][0] == BG_ROUGE_CLAIR_3) {
      sellSum += valuesTransaction[i][0];
      sellQuantity += valuesTransaction[i][1];
    } else {
      break;
    }
  }

  currentPrice = SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_CURRENT_PRICE)
  remainingBet = buySum - sellSum;
  remainingQuantity = buyQuantity - sellQuantity;
  if (buyQuantity != 0) {
    averageBuyingPrice = buySum / buyQuantity
    if (currentPrice < averageBuyingPrice) {
      pnl = -(1 - (currentPrice / averageBuyingPrice));
    } else {
      pnl = (currentPrice / averageBuyingPrice) - 1;
    }
  } else {
    averageBuyingPrice = 0
  }

  valueCrypto = remainingQuantity * currentPrice;
  review = valueCrypto - remainingBet;

  distribution = valueCrypto / valueTotalInvested;

  cell = SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_REMAIN_BET)
  cell.setNumberFormat(getCellFormatNumberDollars(remainingBet));
  cell.setValue(remainingBet);

  cell = SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_REMAIN_QUANTITY);
  cell.setNumberFormat(getCellFormatNumber(remainingQuantity));
  cell.setValue(remainingQuantity);

  cell = SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_MEAN_BUY_PRICE);
  cell.setNumberFormat(getCellFormatNumberDollars(averageBuyingPrice));
  cell.setValue(averageBuyingPrice);

  cell = SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_PNL);
  cell.setFormula("IF(" + sellSum.toFixed(0) + "+" + SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ">" + buySum.toFixed(0) + ";((" + sellSum.toFixed(0) + "+" + SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")-1;-(1-((" + sellSum.toFixed(0) + "+" + SHEET.getRange(rowCryptoCR, CRYPTO_REVIEW.column + CR.NB_COL_VALUE).getA1Notation() + ")/" + buySum.toFixed(0) + ")))")

  descendingSortCR();
}

function majAllSumBuySell() {
  let ui = SpreadsheetApp.getUi();
  const CRYPTO_REVIEW = findCellWithValueAllSheet(CR.VALUE_HEADER);

  if (!CRYPTO_REVIEW) {
    ui.alert('Erreur', "|crDataMAJ| !CRYPTO_REVIEW", ui.ButtonSet.OK);
    return;
  }

  const ROW_FIRST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_BEFORE_FIRST_CRYPTO, CRYPTO_REVIEW.column) + 1;
  const ROW_LAST_CRYPTO_CR = getRowIndexInColumnWithValue(CR.VALUE_ROW_SC, CRYPTO_REVIEW.column);

  const rangeCR = SHEET.getRange(ROW_FIRST_CRYPTO_CR, CRYPTO_REVIEW.column, (ROW_LAST_CRYPTO_CR - ROW_FIRST_CRYPTO_CR), 9);
  const valuesCR = rangeCR.getValues();

  let rangeTransaction;
  let valuesTransaction;
  let indexCryptoTransaction;
  let sellSum;
  let buySum;
  let backgrounds;

  let cryptoSummary = "";

  for (let i = 0; i < valuesCR.length; i++) {
    indexCryptoTransaction = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_TH, valuesCR[i][0], TH.BG_CRYPTO_NAME);
    if (!indexCryptoTransaction) {
      ui.alert('Erreur', "|crDataMAJ| !indexCryptoTransaction | Crypto: " + valuesCR[i][0], ui.ButtonSet.OK);
      continue;
    }

    sellSum = 0;
    buySum = 0;

    rangeTransaction = SHEET_TRANSACTION_HISTORIC.getRange((indexCryptoTransaction.row + 3), indexCryptoTransaction.column, SHEET_TRANSACTION_HISTORIC.getLastRow(), 4);
    valuesTransaction = rangeTransaction.getValues();
    backgrounds = rangeTransaction.getBackgrounds();

    for (let j = 0; j < valuesTransaction.length; j++) {
      if (backgrounds[j][0] == BG_VERT_CLAIR_3) {
        buySum += valuesTransaction[j][0];
      } else if (backgrounds[j][0] == BG_ROUGE_CLAIR_3) {
        sellSum += valuesTransaction[j][0];
      } else {
        break;
      }
    }
    buySum.toFixed(0)
    sellSum.toFixed(0)

    cryptoSummary += valuesCR[i][0] + ":" + buySum.toFixed(0) + "," + sellSum.toFixed(0) + ";"
  }
  cryptoSummary.slice(0, -1)
  SCRIPT_PROPERTIES.setProperty('mapSumBuySell', cryptoSummary.slice(0, -1));
}

function updateMapSumBuySell(cryptoName, transactionAmount, isBuy) {
  let mapSumBuySellString = SCRIPT_PROPERTIES.getProperty('mapSumBuySell') || "";
  let mapSumBuySell = {};

  // Convertir la chaîne en objet
  mapSumBuySellString.split(";").forEach(keyValuePair => {
    let [key, value] = keyValuePair.split(":");
    value = value.split(",");
    mapSumBuySell[key] = { buySum: parseFloat(value[0]), sellSum: parseFloat(value[1]) };
  });

  // Mise à jour des sommes
  if (!mapSumBuySell[cryptoName]) {
    mapSumBuySell[cryptoName] = { buySum: 0, sellSum: 0 };
  }
  if (isBuy) {
    mapSumBuySell[cryptoName].buySum += transactionAmount;
  } else {
    mapSumBuySell[cryptoName].sellSum += transactionAmount;
  }

  // Convertir l'objet en chaîne pour le stockage
  let newMapSumBuySellString = Object.entries(mapSumBuySell).map(([key, val]) => {
    return `${key}:${val.buySum.toFixed(0)},${val.sellSum.toFixed(0)}`;
  }).join(";");

  // Mettre à jour la propriété de script
  SCRIPT_PROPERTIES.setProperty('mapSumBuySell', newMapSumBuySellString);
}

function extractSumBuySellFormula(row, column) {
  const cell = SHEET.getRange(row, column);
  const formula = cell.getFormula();
  const numberPattern = /\b\d+\b/g; 
  const numbers = formula.match(numberPattern);

  if (numbers) {
    return {
      buySum: numbers[1],
      sellSum: numbers[0]
    }
  } else {
    console.log("Aucun nombre spécifique trouvé dans la formule");
    return {
      buySum: 0,
      sellSum: 0
    }
  }
}

function test() {
  extractSumBuySellFormula(5, 13)
}


function callOpenseaApi() {
  const NFT_REVIEW = findCellWithValueAllSheet(NR.VALUE_HEADER);
  let nftInfos;

  if (!NFT_REVIEW) {
    console.log("|callOpenseaApi| !NFT_REVIEW");
    return;
  }
  let cell;
  let ethPrice = getCryptoPriceWithID(ID_ETH_CMC);
  let floorPriceInEth;
  let floorPriceUsd;
  let COLUMN_NR = NFT_REVIEW.column;
  let ROW_FIRST_NFT_NR = getRowIndexInColumnWithValue(NR.VALUE_BEFORE_FIRST_NFT, COLUMN_NR) + 1;

  if (!COLUMN_NR || !ROW_FIRST_NFT_NR) {
    console.log("|callOpenseaApi| !COLUMN_NR || !ROW_FIRST_NFT_NR");
    return;
  }
  const startRow = ROW_FIRST_NFT_NR;
  let floorPriceCache = {}; // Cache for storing floor prices

  for (let row = startRow; row <= LAST_ROW; row++) {
    const nftName = SHEET.getRange(row, COLUMN_NR).getValue();
    if (nftName === NR.VALUE_ROW_TOTAL_INVESTED) {
      break;
    }
    nftInfos = getCollectionNameAndChain(nftName);
    if (nftInfos) {
      // Check if floor price is already in the cache
      if (!floorPriceCache[nftInfos.collectionName]) {
        floorPriceInEth = getFloorPrice(nftInfos.collectionName);
        if (floorPriceInEth) {
          // Store the floor price in the cache
          floorPriceCache[nftInfos.collectionName] = floorPriceInEth;
        }
      } else {
        // Use the cached floor price
        floorPriceInEth = floorPriceCache[nftInfos.collectionName];
      }

      if (floorPriceInEth) {
        cell = SHEET.getRange(row, COLUMN_NR + 3); // cell floor price
        floorPriceUsd = ethPrice * floorPriceInEth;
        cell.setNumberFormat(getCellFormatNumberDollars(floorPriceUsd));
        cell.setValue(floorPriceUsd);
      }
    }
  }
}


function getFloorPrice(collectionName) {
  console.log("..............")
  const apiKeyOpensea = "a5f9fe9ff972462ba0048245f9ec8f54"
  let url = "https://api.opensea.io/api/v1/collection/" + collectionName + "/stats";
  let options = {
    'method' : 'get',
    'headers': {
      'X-API-KEY': apiKeyOpensea
    }
  };

  let response = UrlFetchApp.fetch(url, options);
  let json = JSON.parse(response.getContentText());
  if(response.getResponseCode() === 200){
    let floorPrice = json.stats.floor_price;
    return floorPrice;
  }
  else{
    console.log("|getFloorPrice| Erreur appel api")
    return null
  }
}


function getNFTsOwnedByAccount(accountAddress, chain) {
  const apiKeyOpensea = "a5f9fe9ff972462ba0048245f9ec8f54"
  let url = "https://api.opensea.io/api/v2/chain/" + chain + "/account/" + accountAddress + "/nfts"
  let options = {
    'method': 'get',
    'headers': {
      'X-API-KEY': apiKeyOpensea
    },
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(url, options);
  let json = JSON.parse(response.getContentText());
  
}

function prout(){
  let res = getFloorPrice("side-cards")
}
function getCollectionNameAndChain(nft) {

  let parts = nft.split("#");
  let nftName = parts[0].trim();
  let mapNftInfos = SCRIPT_PROPERTIES.getProperty('mapNftInfos') || "";

  let mapNft = {};

  mapNftInfos.split(";").forEach(keyValuePair => {
    let [key, value] = keyValuePair.split(":");
    value = value.split(",");
    mapNft[key] = { collectionName: value[0], chain: value[1] };
  });

  if (mapNft.hasOwnProperty(nftName)) {
    return { collectionName: mapNft[nftName].collectionName, chain: mapNft[nftName].chain }
  } else {
    console.log("|getCollectionName| "+nft+" n'a pas de collection name associée")
    return null
  }
}

function addToOldCrypto(crypto) {
  let sourceRange = getRangeCrypto(SHEET_NAME_TH, crypto)
  let targetRange
  let cryptoIndex = findCellsWithValueAndBackgroundAllSheet(SHEET_NAME_OLD_CRYPTO, crypto, TH.BG_CRYPTO_NAME)
  if (cryptoIndex) {
    targetRange = SHEET_TRANSACTION_OLD_CRYPTO.getRange(cryptoIndex.row, cryptoIndex.column, sourceRange.getNumRows(), sourceRange.getNumColumns());
    sourceRange.copyTo(targetRange, { contentsOnly: true });
    sourceRange.copyTo(targetRange, { formatOnly: true });
  }
  else {
    SHEET_TRANSACTION_OLD_CRYPTO.insertColumnsBefore(1, 5);
    targetRange = SHEET_TRANSACTION_OLD_CRYPTO.getRange(2, 2, sourceRange.getNumRows(), sourceRange.getNumColumns());
    sourceRange.copyTo(targetRange, { contentsOnly: true });
    sourceRange.copyTo(targetRange, { formatOnly: true });
  }
}


function getRangeCrypto(sheetName, cryptoName) {
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let selectedOptionIndex = findCellsWithValueAndBackgroundAllSheet(sheetName, cryptoName, TH.BG_CRYPTO_NAME);
  if (!selectedOptionIndex) {
    console.log("|getRangeCrypto| " + cryptoName + "n'a pas été trouvé");
    return;
  }

  let values = currentSheet.getRange((selectedOptionIndex.row + TH.NB_ROW_FIRST_TRANSACTION), selectedOptionIndex.column, LAST_ROW, 1).getValues();
  let newTransactionRow;
  for (let j = 0; j < values.length; j++) {
    if (values[j][0] === "") {
      newTransactionRow = j + TH.NB_ROW_FIRST_TRANSACTION + selectedOptionIndex.row - 1;
      break;
    }
  }
  return currentSheet.getRange(selectedOptionIndex.row, selectedOptionIndex.column, (newTransactionRow - selectedOptionIndex.row + 1), TH.NB_COLUMN);
}
