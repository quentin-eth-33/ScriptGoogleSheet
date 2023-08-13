function showFormAddNewTransaction() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let transactionHistoric = findCellWithValueAllSheet("Historique Transaction:");
  let ui = SpreadsheetApp.getUi();

  if (transactionHistoric) {
    let columnTransactionHistoric = transactionHistoric.column;
    console.log("L'index de la colonne columnTransactionHistoric est : " + columnTransactionHistoric);
    let rowBitcoinInTransactionHistoric = getRowIndexInColumnWithValue("Bitcoin", columnTransactionHistoric);

    if (rowBitcoinInTransactionHistoric) {
      console.log("L'index de la ligne rowBitcoinInTransactionHistoric est : " + rowBitcoinInTransactionHistoric);
      let searchRangeCryptoName = "" + getColumnHeader(columnTransactionHistoric) + "" + rowBitcoinInTransactionHistoric + ":" + getColumnHeader(columnTransactionHistoric);
      let lastRow = sheet.getLastRow();
      let dataRange = sheet.getRange(searchRangeCryptoName + lastRow);
      let dataValues = dataRange.getValues();
      let emptyCellTolerance = 20;

      let choices = [];
      let emptyCellCount = 0;

      for (let i = 0; i < dataValues.length && emptyCellCount < emptyCellTolerance; i++) {
        let cellValue = dataValues[i][0];

        if (cellValue !== "") {
          choices.push(cellValue);
          emptyCellCount = 0;
        } else {
          emptyCellCount++;
        }
      }

      let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewTransaction')
        .setWidth(600)
        .setHeight(900);

      htmlOutput.append('<script>let choices = ' + JSON.stringify(choices) + ';</script>');
      ui.showModalDialog(htmlOutput, 'Ajouter une transaction:');
    } else {
      console.log("La valeur Bitcoin n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur Historique Transaction n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Historique Transaction n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function addNewTransaction(selectedAmount, selectedOption, selectedQuantity, selectedDate, isBuy) {

  let ui = SpreadsheetApp.getUi();

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let transactionHistoric = findCellWithValueAllSheet("Historique Transaction:");

  if (transactionHistoric) {
    let columnTransactionHistoric = transactionHistoric.column;
    let nbRowBetweenBuySell = 5;
    let background = "#d9ead3"; // Vert clair 3

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

    console.log("Option sélectionnée : " + selectedOption + " Date : " + selectedDate + " Montant : " + selectedAmount + " Quantité : " + selectedQuantity);

    let columnWithNameOfCrypto = sheet.getRange("" + getColumnHeader(columnTransactionHistoric) + ":" + getColumnHeader(columnTransactionHistoric)).getValues();
    let rowIndex = -1;

    // Recherche de la ligne correspondant à la crypto sélectionnée
    for (let i = 0; i < columnWithNameOfCrypto.length; i++) {
      if (columnWithNameOfCrypto[i][0] === selectedOption) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      let indexRowToMerge = 0;

      if (isBuy === false) {
        rowIndex = rowIndex + nbRowBetweenBuySell;
        background = "#f4cccc"; // Rouge clair 3
        indexRowToMerge = rowIndex - 1; // Ligne ou est présente la ligne grise fusionnée
      }
      else {
        indexRowToMerge = rowIndex + 4; // Ligne ou est présente la ligne grise fusionnée
      }

      // Recherche de la colonne ou insérer la transaction
      let values = sheet.getRange(rowIndex, 1, 4, sheet.getLastColumn()).getValues();
      let columnIndex = -1;
      let columnHeader = "";
      for (let j = columnTransactionHistoric; j < values[0].length; j++) {
        if (values[0][j] === "") {
          columnIndex = j + 1;
          break;
        }
      }

      if (columnIndex !== -1) {
        columnHeader = getColumnHeader(columnIndex);
        console.log("Header de la colonne où est insérée la transaction : " + columnHeader);

        let range = sheet.getRange(rowIndex, columnIndex, 4, 1);
        range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        range.setBackground(background);

        let cell = range.getCell(1, 1); // Première cellule de la plage
        let dateFormat = "dd/MM/YYYY";
        range.setNumberFormat(dateFormat);
        cell.setFontWeight("bold");
        cell.setHorizontalAlignment("right");
        let dateComponents = selectedDate.split("-");
        let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
        cell.setValue(formattedDate);

        cell = range.getCell(2, 1);
        cell.setFontWeight("bold");
        cell.setHorizontalAlignment("right");
        cell.setNumberFormat(getCellFormatNumberDollars(selectedAmount));
        cell.setValue(selectedAmountCell);

        cell = range.getCell(3, 1);
        cell.setFontWeight("bold");
        cell.setHorizontalAlignment("right");
        cell.setNumberFormat(getCellFormatNumber(selectedQuantity));
        cell.setValue(selectedQuantityCell);

        cell = range.getCell(4, 1); // Première cellule de la plage
        cell.setFontWeight("bold");
        cell.setHorizontalAlignment("right");
        let averageBuyingPrice = selectedAmount / selectedQuantity;
        console.log("averageBuyingPrice: " + averageBuyingPrice)
        cell.setNumberFormat(getCellFormatNumberDollars(averageBuyingPrice));
        cell.setValue(averageBuyingPrice);

        let endMergeColumns = getMergedCellColumnsEnd(indexRowToMerge, (columnTransactionHistoric + 1));

        if (endMergeColumns < columnIndex) {
          console.log("endMergeColumns: " + endMergeColumns);
          console.log("columnIndex: " + columnIndex);
          let startRange = sheet.getRange(indexRowToMerge, (columnTransactionHistoric + 1));
          let endRange = sheet.getRange(indexRowToMerge, columnIndex);
          mergeCells(startRange, endRange);
        }

        let cryptoReview = findCellWithValueAllSheet("Bilan Crypto:");
        if (cryptoReview) {
          let columnCryptoReview = cryptoReview.column;
          let rowSelectedOption = getRowIndexInColumnWithValue(selectedOption, columnCryptoReview);

          if (rowSelectedOption) {
            cell = sheet.getRange(rowSelectedOption, (columnCryptoReview + 2));
            cell.setNumberFormat(getCellFormatNumber(cell.getValue()));
            cell = sheet.getRange(rowSelectedOption, (columnCryptoReview + 3));
            cell.setNumberFormat(getCellFormatNumberDollars(cell.getValue()));
          }
          else {
            console.log("La rowSelectedOption n'a pas été trouvée sur la feuille.");
            ui.alert('Erreur', "La rowSelectedOption n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
          }

        }
        else {
          console.log("La valeur Bilan Crypto n'a pas été trouvée sur la feuille.");
          ui.alert('Erreur', "La valeur Bilan Crypto n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
        }
      } else {
        console.log("Pas assez de colonne pour pouvoir insérer une transaction");
        ui.alert('Erreur', "Pas assez de colonne pour pouvoir insérer une transaction", ui.ButtonSet.OK);
      }
    } else {
      console.log(selectedOption + "n'a pas été trouvé");
      ui.alert('Erreur', selectedOption + "n'a pas été trouvé", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur Historique Transaction n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Historique Transaction n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}


function callCmcApi() {
  let ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  let mapIdCryptoGlobalString = scriptProperties.getProperty('mapIdCryptoGlobal');

  // Convertir la chaîne en map
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    mapIdCryptoGlobal[key] = value;
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cryptoReview = findCellWithValueAllSheet("Bilan Crypto:");

  if (cryptoReview) {
    let columnCryptoReview = cryptoReview.column;
    console.log("L'index de la columnCryptoReview est : " + columnCryptoReview);

    let rowBitcoinCryptoReview = getRowIndexInColumnWithValue("Bitcoin", columnCryptoReview)

    if (rowBitcoinCryptoReview) {
      console.log("L'index de la rowBitcoinCryptoReview est : " + rowBitcoinCryptoReview);
      const startRow = rowBitcoinCryptoReview;
      const endRow = sheet.getLastRow();
      const urlSymbols = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
      const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
      const apiKeyList = getApiKeyCmcList();
      let indexTabApiKey = 0;
      const indexColumnCurrentPrice = columnCryptoReview + 4;
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
        const cryptoName = sheet.getRange(row, columnCryptoReview).getValue();
        isCallValid = false;
        if (cryptoName === "StableCoin Binance") {
          console.log("Fin de la récupération des noms de cryptomonnaies.");
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
            scriptProperties.setProperty('mapIdCryptoGlobal', mapIdCryptoGlobalString);
            isCallValid = true;
          } else if (response.getResponseCode() === 429) {
            if (indexTabApiKey < apiKeyList.length - 1) {
              indexTabApiKey = indexTabApiKey + 1;
              options.headers["X-CMC_PRO_API_KEY"] = apiKeyList[indexTabApiKey];
              row = row - 1;
            } else {
              console.log("plus d'api key opérationnelle disponible")
              ui.alert('Erreur', "plus d'api key opérationnelle disponible", ui.ButtonSet.OK);
            }

          } else {
            console.log(`Erreur lors de la récupération des symboles de cryptomonnaie : ${response.getResponseCode()} - ${jsonData}`);
            ui.alert('Erreur', "Erreur lors de la récupération des symboles de cryptomonnaie", ui.ButtonSet.OK);
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
      console.log(idsAsString);

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
                sheet.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setNumberFormat(getCellFormatNumberDollars(cryptoPrice));
                sheet.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setFontWeight("bold");
                sheet.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setHorizontalAlignment("right");
                sheet.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setBackground("#f3f3f3");
                sheet.getRange(cryptosInfo[crypto].indexRow, indexColumnCurrentPrice).setValue(cryptoPrice);
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
          ui.alert('Erreur', "plus d'api key opérationnelle disponible", ui.ButtonSet.OK);
        }
      }
      else {
        console.log(`Erreur lors de la récupération des données de la cryptomonnaie '${cryptoName}' : ${quoteResponse.getResponseCode()} - ${quoteJsonData}`);
        ui.alert('Erreur', "Erreur lors de la récupération des données de la cryptomonnaie", ui.ButtonSet.OK);
      }
    } else {
      console.log("La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.");
      ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la colonne crypto review.", ui.ButtonSet.OK);
    }

    console.log("mapIdCryptoGlobalString finale: " + mapIdCryptoGlobalString);

  } else {
    console.log("La valeur Bilan Crypto n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Bilan Crypto n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}


function showFormAddNewNFT() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let transactionHistoric = findCellWithValueAllSheet("Historique Transaction:");
  let ui = SpreadsheetApp.getUi();

  if (transactionHistoric) {
    let columnTransactionHistoric = transactionHistoric.column;
    console.log("L'index de la colonne columnTransactionHistoric est : " + columnTransactionHistoric);
    let rowBitcoinInTransactionHistoric = getRowIndexInColumnWithValue("Bitcoin", columnTransactionHistoric);

    if (rowBitcoinInTransactionHistoric) {
      console.log("L'index de la ligne rowBitcoinInTransactionHistoric est : " + rowBitcoinInTransactionHistoric);
      let searchRangeCryptoName = "" + getColumnHeader(columnTransactionHistoric) + "" + rowBitcoinInTransactionHistoric + ":" + getColumnHeader(columnTransactionHistoric);
      let lastRow = sheet.getLastRow();
      let dataRange = sheet.getRange(searchRangeCryptoName + lastRow);
      let dataValues = dataRange.getValues();
      let emptyCellTolerance = 20;

      let choices = [];
      let emptyCellCount = 0;

      for (let i = 0; i < dataValues.length && emptyCellCount < emptyCellTolerance; i++) {
        let cellValue = dataValues[i][0];

        if (cellValue !== "") {
          choices.push(cellValue);
          emptyCellCount = 0;
        } else {
          emptyCellCount++;
        }
      }

      let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewNFT')
        .setWidth(600)
        .setHeight(900);

      htmlOutput.append('<script>let choices = ' + JSON.stringify(choices) + ';</script>');
      ui.showModalDialog(htmlOutput, 'Nouvel Achat NFT:');
    } else {
      console.log("La valeur Bitcoin n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }
  } else {
    console.log("La valeur Historique Transaction n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Historique Transaction n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function addNewNFT(idNftInput, optionSelectSell, amountInput, quantityInputSell, selectedDate) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let ui = SpreadsheetApp.getUi();
  console.log("Id NFT: " + idNftInput + " | optionSelectSell: " + optionSelectSell + " | amountInput: " + amountInput + " | quantityInputSell: " + quantityInputSell)

  let amountInputCell = amountInput;
  if (amountInput.indexOf('.') !== -1) {
    amountInputCell = amountInput.replace('.', ',');
  } else if (amountInput.indexOf(',') !== -1) {
    amountInput = amountInput.replace(',', '.');
  }

  let nftReview = findCellWithValueAllSheet("Bilan NFT:");
  if (nftReview) {
    let columnNftReview = nftReview.column;
    let searchValue = "Total Investi:";
    rowSearchValue = getRowIndexInColumnWithValue(searchValue, columnNftReview);
    if (rowSearchValue) {
      console.log("Index de la crypto cherchée: " + rowSearchValue);

      sheet.getRange(rowSearchValue, columnNftReview, 2, 4).moveTo(sheet.getRange((rowSearchValue + 1), columnNftReview));

      let cell = sheet.getRange(rowSearchValue, columnNftReview, 1, 4);
      cell.setBackground("#f3f3f3");
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("right");
      cell.setFontSize(8);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      cell = sheet.getRange(rowSearchValue, columnNftReview);
      cell.setBackground("#cfe2f3");
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("left");
      cell.setFontSize(8);
      cell.setFontColor("#4285f4");
      cell.setValue(idNftInput);

      cell = sheet.getRange(rowSearchValue, (columnNftReview + 1));
      cell.setNumberFormat(getCellFormatNumberDollars(amountInput));
      cell.setValue(amountInputCell);

      cell = sheet.getRange(rowSearchValue, (columnNftReview + 2));
      cell.setValue("Non défini");

      cell = sheet.getRange(rowSearchValue, (columnNftReview + 3));
      cell.setNumberFormat("0$");
      cell.setValue(0);

      addNewTransaction(amountInput, optionSelectSell, quantityInputSell, selectedDate, false);
    } else {
      console.log("La valeur Total investi n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur Total investi n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }

  } else {
    console.log("La valeur Bilan NFT n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Bilan NFT n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }

}


function showFormAddNewCrypto() {
  let ui = SpreadsheetApp.getUi();
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCrypto')
    .setWidth(300)
    .setHeight(220);

  ui.showModalDialog(htmlOutput, 'Ajouter une nouvelle crypto:');
}

function addNewCrypto(cryptoName) {
  let ui = SpreadsheetApp.getUi();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  console.log("Nom de la crypto à ajouter: " + cryptoName);
  let searchValue = "StableCoin Binance"; // A custom
  let transactionHistoric = findCellWithValueAllSheet("Historique Transaction:");

  if (transactionHistoric) {
    let columnTransactionHistoric = transactionHistoric.column;
    console.log("columnTransactionHistoric: " + columnTransactionHistoric);

    let indexSearchValue = getRowIndexInColumnWithValue(searchValue, columnTransactionHistoric);

    if (indexSearchValue) {
      console.log("Index de la crypto cherchée: " + indexSearchValue);
      let buyTab = ["Date", "Montant", "Quantité", "Prix d'Achat"];
      let sellTab = ["Date", "Montant", "Quantité", "Prix de Vente"];

      sheet.insertRowsBefore(indexSearchValue - 1, 10); // -1 -> car sinon ca prend le format de "StableCoin Bincance" (cellule colorées etc)

      console.log("indexSearchValue: ", indexSearchValue)
      console.log("columnTransactionHistoric: ", columnTransactionHistoric)
      let startMergingRange = sheet.getRange(indexSearchValue, columnTransactionHistoric); // Cellule de début de la fusion
      let endMergingRange = sheet.getRange((indexSearchValue + 8), columnTransactionHistoric);   // Cellule de fin de la fusion
      mergeCells(startMergingRange, endMergingRange);

      let cell = sheet.getRange(indexSearchValue, columnTransactionHistoric);
      //#ebcfff --> background violet
      cell.setBackground("#ebcfff");
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle"); // Alignement vertical au centre
      cell.setFontSize(15); // Taille de police 15
      cell.setFontColor("#9900ff");
      cell.setValue(cryptoName);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      startMergingRange = sheet.getRange(indexSearchValue, (columnTransactionHistoric + 1)); // Cellule de début de la fusion
      endMergingRange = sheet.getRange((indexSearchValue + 3), (columnTransactionHistoric + 1));   // Cellule de fin de la fusion
      mergeCells(startMergingRange, endMergingRange);

      cell = sheet.getRange(indexSearchValue, (columnTransactionHistoric + 1));
      cell.setBackground("#fce5cd");
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle"); // Alignement vertical au centre
      cell.setFontSize(10);
      cell.setFontColor("#ff9900");
      cell.setValue("ACHAT");
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      startMergingRange = sheet.getRange((indexSearchValue + 5), (columnTransactionHistoric + 1)); // Cellule de début de la fusion
      endMergingRange = sheet.getRange((indexSearchValue + 8), (columnTransactionHistoric + 1));   // Cellule de fin de la fusion
      mergeCells(startMergingRange, endMergingRange);

      cell = sheet.getRange((indexSearchValue + 5), (columnTransactionHistoric + 1));
      cell.setBackground("#fce5cd");
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle"); // Alignement vertical au centre
      cell.setFontSize(10);
      cell.setFontColor("#ff9900");
      cell.setValue("VENTE");
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      // couleur gris: #cccccc
      startMergingRange = sheet.getRange((indexSearchValue + 4), (columnTransactionHistoric + 1)); // Cellule de début de la fusion
      endMergingRange = sheet.getRange((indexSearchValue + 4), (columnTransactionHistoric + 2));   // Cellule de fin de la fusion
      let mergingRange = sheet.getRange(startMergingRange.getRow(), startMergingRange.getColumn(), 1, 2); // Car la fonction permet de fusionner verticalement et non horizontalement
      mergingRange.merge();
      cell = sheet.getRange((indexSearchValue + 4), (columnTransactionHistoric + 1));
      cell.setBackground("#cccccc");
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      for (let i = 0; i < buyTab.length; i++) {
        setBleueTexte((indexSearchValue + i), (columnTransactionHistoric + 2), buyTab[i]);
      }

      for (let i = 0; i < sellTab.length; i++) {
        setBleueTexte((indexSearchValue + i + 5), (columnTransactionHistoric + 2), sellTab[i]);
      }

      let cryptoReview = findCellWithValueAllSheet("Bilan Crypto:"); // Customable
      if (cryptoReview) {
        let columnCryptoReview = cryptoReview.column;
        let rowSearchValue = getRowIndexInColumnWithValue("StableCoin Binance", columnCryptoReview);
        if (rowSearchValue) {
          let rowReview = getRowIndexInColumnWithValue("Bilan:", columnCryptoReview);
          if (rowReview) {
            sheet.getRange(rowSearchValue, columnCryptoReview, (rowReview - rowSearchValue + 1), 9).moveTo(sheet.getRange((rowSearchValue + 1), columnCryptoReview));

            cell = sheet.getRange(rowSearchValue, columnCryptoReview);
            cell.setBackground("#cfe2f3");
            cell.setFontWeight("bold");
            cell.setHorizontalAlignment("left");
            cell.setFontSize(8);
            cell.setFontColor("#4285f4");
            cell.setValue(cryptoName);
            cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 1), 1, 8);
            cell.setBackground("#f3f3f3");
            cell.setFontWeight("bold");
            cell.setHorizontalAlignment("right");
            cell.setFontSize(8);
            cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            let firstColumnSum = getColumnHeader(columnTransactionHistoric + 3);
            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 1));
            cell.setFormula("=SUM(" + firstColumnSum + "" + (indexSearchValue + 1) + ":" + (indexSearchValue + 1) + ")-SUM(" + firstColumnSum + "" + (indexSearchValue + 6) + ":" + (indexSearchValue + 6) + ")");
            cell.setNumberFormat('0$');

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 2));
            cell.setFormula("=SUM(" + firstColumnSum + "" + (indexSearchValue + 2) + ":" + (indexSearchValue + 2) + ")-SUM(" + firstColumnSum + "" + (indexSearchValue + 7) + ":" + (indexSearchValue + 7) + ")");

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 3));
            cell.setFormula("=IF(SUM(" + firstColumnSum + "" + (indexSearchValue + 2) + ":" + (indexSearchValue + 2) + ") =0; 0; SUM(" + firstColumnSum + "" + (indexSearchValue + 1) + ":" + (indexSearchValue + 1) + ")/SUM(" + firstColumnSum + "" + (indexSearchValue + 2) + ":" + (indexSearchValue + 2) + "))");
            cell.setNumberFormat('0$');

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 4));
            cell.setValue(0);
            cell.setNumberFormat('0$');

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 5));
            cell.setFormula("=" + getColumnHeader(columnCryptoReview + 4) + "" + rowSearchValue + "*" + getColumnHeader(columnCryptoReview + 3) + "" + rowSearchValue);
            cell.setNumberFormat('0$');

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 6));
            cell.setFormula("=" + getColumnHeader(columnCryptoReview + 5) + "" + rowSearchValue + "-" + getColumnHeader(columnCryptoReview + 1) + "" + rowSearchValue);
            cell.setNumberFormat('[Color50]+0$;[RED]-0$');

            cell = sheet.getRange("L" + rowSearchValue);
            cell.setFormula("=IF(" + getColumnHeader(columnCryptoReview + 3) + "" + rowSearchValue + "=0;0;IF(" + getColumnHeader(columnCryptoReview + 4) + "" + rowSearchValue + "<" + getColumnHeader(columnCryptoReview + 3) + "" + rowSearchValue + ";-(1-(" + getColumnHeader(columnCryptoReview + 4) + "" + rowSearchValue + "/" + getColumnHeader(columnCryptoReview + 3) + "" + rowSearchValue + "));(" + getColumnHeader(columnCryptoReview + 4) + "" + rowSearchValue + "/" + getColumnHeader(columnCryptoReview + 3) + "" + rowSearchValue + ")-1))");
            cell.setNumberFormat('[Color50]+0.00%;[Red]-0.00%');

            cell = sheet.getRange(rowSearchValue, (columnCryptoReview + 8));
            cell.setFormula("=" + getColumnHeader(columnCryptoReview + 5) + "" + rowSearchValue + "/" + getColumnHeader(columnCryptoReview + 1) + "" + (rowSearchValue + 6))
            cell.setNumberFormat('0.00%');

            let rowBitcoinCryptoReview = getRowIndexInColumnWithValue("Bitcoin", columnCryptoReview);
            if (rowBitcoinCryptoReview) {
              // Créer la règle de mise en forme conditionnelle
              let plageFormatConditionnelle = sheet.getRange("" + getColumnHeader(columnCryptoReview + 1) + "" + rowBitcoinCryptoReview + ":" + getColumnHeader(columnCryptoReview + 8) + "" + rowSearchValue);
              let regleMiseEnForme = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied("=$" + getColumnHeader(columnCryptoReview + 1) + "" + rowBitcoinCryptoReview + "<=0")
                .setBackground("#d9ead3") // Couleur de fond verte en cas de condition satisfaite
                .setRanges([plageFormatConditionnelle])
                .build();
              sheet.setConditionalFormatRules([regleMiseEnForme]);
            } else {
              console.log("La valeur Bitcoin n'a pas été trouvée sur la feuille.");
              ui.alert('Erreur', "La valeur Bitcoin n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
            }

            let rowTotalAssets = getRowIndexInColumnWithValue("Total Actif:", columnCryptoReview);
            if (rowTotalAssets) {
              cell = sheet.getRange(rowTotalAssets, (columnCryptoReview + 1));
              cell.setFormula("=SUM(" + getColumnHeader(columnCryptoReview + 5) + "" + rowBitcoinCryptoReview + ":" + getColumnHeader(columnCryptoReview + 5) + "" + rowSearchValue + ";" + getColumnHeader(columnCryptoReview + 1) + "" + (rowSearchValue + 1) + ":" + getColumnHeader(columnCryptoReview + 1) + "" + (rowSearchValue + 5) + ") - SUMIF(" + getColumnHeader(columnCryptoReview + 1) + "" + rowBitcoinCryptoReview + ":" + getColumnHeader(columnCryptoReview + 1) + "" + rowSearchValue + "; \"<0\")");

            } else {
              console.log("La valeur Total Actif n'a pas été trouvée sur la feuille.");
              ui.alert('Erreur', "La valeur Total Actif n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
            }

          } else {
            console.log("La valeur Bilan n'a pas été trouvée sur la feuille.");
            ui.alert('Erreur', "La valeur Bilan n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
          }

        } else {
          console.log("La valeur StableCoin Binance n'a pas été trouvée sur la feuille.");
          ui.alert('Erreur', "La valeur StableCoin Binance n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
        }

      }
      else {
        console.log("La valeur Bilan Crypto n'a pas été trouvée sur la feuille.");
        ui.alert('Erreur', "La valeur Bilan Crypto n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
      }
    } else {
      console.log("La valeur indexSearchValue n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur indexSearchValue n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }

  } else {
    console.log("La valeur Historique Transaction n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur Historique Transaction n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function showFormAddNewCashIn() {
  let ui = SpreadsheetApp.getUi();
  let htmlOutput = HtmlService.createHtmlOutputFromFile('formAddNewCashIn')
    .setWidth(300)
    .setHeight(330);

  ui.showModalDialog(htmlOutput, 'New Cash In:');
}

function addNewCashIn(amount, date) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let ui = SpreadsheetApp.getUi();
  let background = "#f3f3f3";

  let totalInvested = findCellWithValueAllSheet("Historique Total Investi:");
  if (totalInvested) {
    let columnTotalInvested = totalInvested.column;
    let rowTotal = getRowIndexInColumnWithValue("Total:", columnTotalInvested);
    if (rowTotal) {
      sheet.getRange(rowTotal, columnTotalInvested, 1, 2).moveTo(sheet.getRange((rowTotal + 1), columnTotalInvested));

      let cell = sheet.getRange(rowTotal, columnTotalInvested);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
      cell.setBackground(background);
      let dateFormat = "dd/MM/YYYY";
      cell.setNumberFormat(dateFormat);
      cell.setFontWeight("bold");
      cell.setFontSize(8);
      cell.setHorizontalAlignment("right");
      let dateComponents = date.split("-");
      let formattedDate = dateComponents[2] + "/" + dateComponents[1] + "/" + dateComponents[0];
      cell.setValue(formattedDate);

      cell = sheet.getRange(rowTotal, (columnTotalInvested + 1))
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
      cell.setBackground(background);
      cell.setFontWeight("bold");
      cell.setFontSize(8);
      cell.setNumberFormat(getCellFormatNumberDollars(amount));
      cell.setValue(amount);

      addNewTransaction(amount, "Euros", amount, date, true);

    } else {
      console.log("La valeur \"Total:\" n'a pas été trouvée sur la feuille.");
      ui.alert('Erreur', "La valeur \"Total:\" n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
    }

  } else {
    console.log("La valeur \"Historique Total Investi:\" n'a pas été trouvée sur la feuille.");
    ui.alert('Erreur', "La valeur \"Historique Total Investi:\" n'a pas été trouvée sur la feuille.", ui.ButtonSet.OK);
  }
}

function addNewChronologicalTransactionHistoric(selectedAmount, selectedOptionBuy, selectedQuantityBuy, selectedOptionSell, selectedQuantitySell, selectedDate) {

  // 1: selectedAmountCell, selectedQuantityCell, averageBuyingPrice, formattedDate
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let ui = SpreadsheetApp.getUi();

  let selectedQuantityBuyCalcul = getFormatCalculationScript(selectedQuantityBuy);
  let selectedAmountCalcul = getFormatCalculationScript(selectedAmount);
  let selectedQuantitySellCalcul = getFormatCalculationScript(selectedQuantitySell);


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

      cell = sheet.getRange(rowLastTransaction, columnChronologicalTransactionHistoric, 1, 8);
      cell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
      cell.setFontWeight("bold");
      cell.setHorizontalAlignment("right");
      cell.setFontSize(8);

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 1), 1, 3);
      cell.setBackground("#d9ead3");

      cell = sheet.getRange(rowLastTransaction, (columnChronologicalTransactionHistoric + 4), 1, 3);
      cell.setBackground("#f4cccc");

      cell = sheet.getRange(rowLastTransaction, columnChronologicalTransactionHistoric);
      cell.setBackground("#cfe2f3")
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
      cell.setBackground("#cfe2f3")
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
