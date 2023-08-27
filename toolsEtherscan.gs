function addTransactionsEtherscan() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let apiKeyEtherscan = "5DDYRKH9WW2CASRHPBR2RZ29467S31V18D"
  let addressses = ["0x914d449A0d989F2CCF52975ced45F925828c8030"/*, "0x897cbe1b142eA9b34A82A7302c84AD19F73A1C70"*/]
  let listSellTransactions;
  let listBuyTransactions;
  const cryptoNameInGS = getAllCryptoNameInGS();

  for (address of addressses) {
    listSellTransactions = [];
    listBuyTransactions = [];

    let listTransactionsToAdd = fetchDataFromEtherscan(address, "etherscan.io");
    if (listTransactionsToAdd) {
      console.log("allTransaction.length: " + listTransactionsToAdd.length);
      for (transactionToAdd of listTransactionsToAdd) {
        console.log("|main| Ajout de la transaction: " + transactionToAdd.hash + " , dans la feuille google sheet");
        if (!(cryptoNameInGS.includes(transactionToAdd.cryptoBuy))) {
          console.log("|main| Ajout de: " + transactionToAdd.cryptoBuy + " à la feuille google sheet")
          //addNewCrypto(transactionToAdd.cryptoBuy);
          cryptoNameInGS.push(transactionToAdd.cryptoBuy);
        }
        if (!(cryptoNameInGS.includes(transactionToAdd.cryptoSell))) {
          console.log("|main| Ajout de: " + transactionToAdd.cryptoSell + " à la feuille google sheet")
          //addNewCrypto(transactionToAdd.cryptoSell);
          cryptoNameInGS.push(transactionToAdd.cryptoSell);
        }
        //transaction(transactionToAdd.amountTransaction, transactionToAdd.cryptoBuy, transactionToAdd.quantityBuy, transactionToAdd.cryptoSell, transactionToAdd.quantitySell, transactionToAdd.date)
      }
    }
  }
}

function getFormattedDate(timestampInSeconds) {
  let timestampInMilliseconds = timestampInSeconds * 1000;
  let dateObj = new Date(timestampInMilliseconds);
  let year = dateObj.getFullYear();
  let month = (dateObj.getMonth() + 1).toString().padStart(2, '0'); // Le mois commence à 0, donc ajoutez 1
  let day = dateObj.getDate().toString().padStart(2, '0');
  let formattedDate = year + '-' + month + '-' + day;
  return formattedDate;
}

function hexToDecimal(hexValue) {
  let decimalValue = parseInt(hexValue, 16);
  return decimalValue;
}

function getLastTimeStampTransaction() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastTimeStampTransaction = scriptProperties.getProperty('lastTimeStampTransaction');
  return parseInt(lastTimeStampTransaction)
}

function setLastTimeStampTransaction(value) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("lastTimeStampTransaction", value);
}

function getLastBlockTransaction() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastTimeStampTransaction = scriptProperties.getProperty('lastBlockTransaction');
  return parseInt(lastTimeStampTransaction)
}

function setLastBlockTransaction(value) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("lastBlockTransaction", value);
}

function fetchDataFromEtherscan(address, apiUse) {

  let lastBlockTransaction = getLastBlockTransaction();
  lastBlockTransaction = 0;
  const url = "https://api." + apiUse + "/api?module=account&action=txlist&address=" + address + "&startblock=" + lastBlockTransaction + "&endblock=99999999&sort=desc&apikey=5DDYRKH9WW2CASRHPBR2RZ29467S31V18D";

  const options = {
    method: "get",
    muteHttpExceptions: true
  };
  const transferTopic = "0xddf252ad1be2c89b69c2b068fc378daa952ba7f163c4a11628f55a4df523b3ef";
  const methodIdSwapExactETHForTokens = "0x7ff36ab5";
  const methodIdSwapExactTokensForETH = "0x18cbafe5";
  const methodIdUnibotBuyV2 = "0x19948479";
  const methodIdUnibotSellV2 = "0x8ee938a9";
  const methodIdUnibotSellV3 = "0x8107aee3";
  const methodIdswapExactETHForTokensSupportingFeeOnTransferTokens ="0xb6f9de95";
  const methodIdMulticall = "0x5ae401dc";
  const methodIdSwapExactTokensForTokens = "0x38ed1739";
  const methodIdSwapExactTokensForETHSupportingFeeOnTransferTokens = "0x791ac947";
  const methodIdExecute = "0x3593564c"
  
  const listMethodIdSwap = [methodIdSwapExactETHForTokens, methodIdSwapExactTokensForETH, methodIdUnibotBuyV2, methodIdUnibotSellV2, methodIdUnibotSellV3, methodIdswapExactETHForTokensSupportingFeeOnTransferTokens, methodIdMulticall, methodIdSwapExactTokensForTokens, methodIdSwapExactTokensForETHSupportingFeeOnTransferTokens, methodIdExecute]

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const lastTimeStampTransaction = /*getLastTimeStampTransaction()*/0;
  let listBuyTransactions = [];
  let listSellTransactions = [];
  let timestampMax = 0;
  let blockNumberMax = 0;
  let transactionReceipt;
  let listTransferLogs;

  let amountTransaction;
  let quantitySell;
  let quantityBuy;
  let cryptoSell;
  let cryptoBuy;
  let date;
  let logs;
  let gasFee;
  let listTransactionsToAdd = [];
  if (data.status === "1" && data.message === "OK") {

    const transactions = data.result;
    for (const transaction of transactions) {
      listTransferLogs = [];
      if (transaction.timeStamp > timestampMax) {
        timestampMax = transaction.timeStamp;
      }
      if (transaction.blockNumber > blockNumberMax) {
        blockNumberMax = transaction.blockNumber;
      }

      if (listMethodIdSwap.includes(transaction.methodId) && transaction.timeStamp > lastTimeStampTransaction && transaction.txreceipt_status == 1) {
        transactionReceipt = getTransactionReceipt(transaction.hash);
        logs = transactionReceipt.logs;
        date = getFormattedDate(transaction.timeStamp);
        gasFee = transaction.gasUsed * (transaction.gasPrice / 10e17)
        for (const log of logs) {
          if (log.topics[0] == transferTopic) {
            listTransferLogs.push(log);
            console.log(log)
          }
        }
        if (listTransferLogs.length == 2) {
          //console.log("Swap classique");
          cryptoSell = getCryptpoName(listTransferLogs[0].address);
          quantitySell = hexToDecimal(listTransferLogs[0].data) / 10e17;
          cryptoBuy = getCryptpoName(listTransferLogs[1].address);
          quantityBuy = hexToDecimal(listTransferLogs[1].data) / 10e17;
          if (cryptoSell == "Ethereum") {
            amountTransaction = (quantitySell + gasFee) * getCryptoPriceWithName(cryptoSell);
          }
          else if (cryptoBuy == "Ethereum") {
            amountTransaction = (quantityBuy + gasFee) * getCryptoPriceWithName(cryptoBuy);
          }
          else {
            console.log("Pas d'eth dans la transaction")
          }
          if (cryptoSell && quantitySell && cryptoBuy && quantityBuy && transaction.hash && amountTransaction && date) {
            listTransactionsToAdd.push({ hash: transaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          }
          else {
            console.log("Une des valeurs de la transaction est undefine")
          }
          showTransactionDetails("2 transfer", transaction.hash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);
        }
        else if (listTransferLogs.length > 2) {
          //console.log("Sa taxe sa race fait gaffe mon canard");
          cryptoSell = getCryptpoName(listTransferLogs[0].address)
          cryptoBuy = getCryptpoName(listTransferLogs[listTransferLogs.length - 1].address)
          if (cryptoSell == "Ethereum" || cryptoBuy == "Ethereum") {
            if (cryptoBuy == "Ethereum") {
              quantitySell = (hexToDecimal(listTransferLogs[0].data) / 10e17) + (hexToDecimal(listTransferLogs[1].data) / 10e17);
              quantityBuy = hexToDecimal(listTransferLogs[listTransferLogs.length - 1].data) / 10e17;
              amountTransaction = (quantityBuy + gasFee) * getCryptoPriceWithName(cryptoBuy);
            } else {
              quantitySell = hexToDecimal(listTransferLogs[0].data) / 10e17;
              quantityBuy = hexToDecimal(listTransferLogs[listTransferLogs.length - 1].data) / 10e17;
              amountTransaction = (quantitySell + gasFee) * getCryptoPriceWithName(cryptoSell);
            }
            listTransactionsToAdd.push({ hash: transaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          }
          showTransactionDetails("taxe", transaction.hash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);
        }
        else {
          console.log("1 seul transfer c'est chelou de zinzin");
        }
      }

    }
    //setLastTimeStampTransaction(timestampMax);
    return listTransactionsToAdd;
  } else {
    console.log("|fetchDataFromEtherscan| API Etherscan message: " + data.message);
    return null;
  }
}

function getTransactionReceipt(hash) {
  let url = "https://api.etherscan.io/api?module=proxy&action=eth_getTransactionReceipt&txhash=" + hash + "&apikey=5DDYRKH9WW2CASRHPBR2RZ29467S31V18D"
  let response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() === 200) {
    let responseData = response.getContentText();
    let jsonData = JSON.parse(responseData);
    return jsonData.result;
  } else {
    console.log("|getTransactionReceipt| Fail request etherscan api, response.getResponseCode(): " + response.getResponseCode());
    return null;
  }
}

function getCryptpoName(address) {
  let name;
  const scriptProperties = PropertiesService.getScriptProperties();
  let mapAddressCryptoGlobalString = scriptProperties.getProperty('mapAddressCryptoGlobal');
  const mapAddressCryptoGlobal = {};
  const keyValuePairs = mapAddressCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    if (key) {
      mapAddressCryptoGlobal[key] = value;
    }

  });
  if (!(address in mapAddressCryptoGlobal)) {
    const url = "https://pro-api.coinmarketcap.com/v2/cryptocurrency/info"
    let apiKey = "ca39beee-0d9a-4bc9-8847-a60f190fc9ad"
    let options = {
      method: "get",
      headers: {
        "X-CMC_PRO_API_KEY": apiKey,
      },
      muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
    };
    let infoResponse = UrlFetchApp.fetch(`${url}?address=${address}`, options);

    if (infoResponse.getResponseCode() === 200) {
      let infoJsonData = JSON.parse(infoResponse.getContentText());
      let firstKey = Object.keys(infoJsonData.data)[0];
      name = infoJsonData.data[firstKey].name;
      if (name) {
        mapAddressCryptoGlobalString += `;${address}:${name}`;
        scriptProperties.setProperty('mapAddressCryptoGlobal', mapAddressCryptoGlobalString);
        if (name == "WETH") {
          name = "Ethereum";
        }
        return name;
      } else {
        console.log("|getCryptpoName| name is null premier else");
      }
    } else {
      console.log("|getCryptpoName| Fail request cmc api, infoResponse.getResponseCode(): " + infoResponse.getResponseCode());
    }
  }
  else {
    name = mapAddressCryptoGlobal[address];
    if (name) {
      if (name == "WETH") {
        name = "Ethereum";
      }
      return name;
    } else {
      console.log("|getCryptpoName| name is null dernier else");
    }
  }
  return null;
}

function getCryptoPriceWithName(name) {
  const urlSymbols = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
  const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  let cryptoId;
  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": "ca39beee-0d9a-4bc9-8847-a60f190fc9ad",
    },
    muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
  };
  let cryptoPrice;
  const scriptProperties = PropertiesService.getScriptProperties();
  let mapIdCryptoGlobalString = scriptProperties.getProperty('mapIdCryptoGlobal');
  const mapIdCryptoGlobal = {};
  const keyValuePairs = mapIdCryptoGlobalString.split(";");
  keyValuePairs.forEach(keyValuePair => {
    const [key, value] = keyValuePair.split(":");
    if (key) {
      mapIdCryptoGlobal[key] = value;
    }
  });

  if (!(name in mapIdCryptoGlobal)) {
    let response = UrlFetchApp.fetch(urlSymbols, options);
    if (response.getResponseCode() === 200) {
      let jsonData = response.getContentText();
      const json = JSON.parse(jsonData);
      const data = json.data;
      const cryptoData = data.find(item => item.name === name);

      if (!cryptoData) {
        console.log(`Cryptomonnaie '${name}' introuvable.`);
      }

      cryptoId = cryptoData.id;
      mapIdCryptoGlobalString += `;${name}:${cryptoId}`;
      scriptProperties.setProperty('mapIdCryptoGlobal', mapIdCryptoGlobalString);
    } else if (response.getResponseCode() === 429) {
      console.log("|getCryptoPriceWithName| nb limite de requete atteint")
      return null;
    } else {
      console.log("|getCryptoPriceWithName| Erreur chelou pour l'appel")
      return null;
    }
  }
  else {
    cryptoId = mapIdCryptoGlobal[name];
    if (!cryptoId) {
      return null;
    }
  }

  let quoteResponse = UrlFetchApp.fetch(`${urlQuotes}?id=${cryptoId}`, options);


  if (quoteResponse.getResponseCode() === 200) {
    let quoteJsonData = quoteResponse.getContentText();
    const quoteJson = JSON.parse(quoteJsonData);

    cryptoPrice = quoteJson.data[cryptoId].quote.USD.price;
    if (!cryptoPrice) {
      return null;
    }
  } else if (quoteResponse.getResponseCode() === 429) {
    console.log("|getCryptoPriceWithName| nb limite de requete atteint")
  } else {
    console.log("|getCryptoPriceWithName| Erreur chelou pour l'appel")
    return null;
  }
  return cryptoPrice;
}

function getAllCryptoNameInGS() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let transactionHistoric = findCellWithValueAllSheet("Historique Transaction:");

  if (transactionHistoric) {
    let columnTransactionHistoric = transactionHistoric.column;
    let rowBitcoinInTransactionHistoric = getRowIndexInColumnWithValue("Bitcoin", columnTransactionHistoric);

    if (rowBitcoinInTransactionHistoric) {
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
      return choices;
    }
    else {
      console.log("|getAllCryptoInGS| " + rowBitcoinInTransactionHistoric + " not found")
    }
  } else {
    console.log("|getAllCryptoInGS| " + transactionHistoric + " not found")
  }
  return null
}

function showTransactionDetails(typeTransaction, transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date) {
  console.log("typeTransaction: " + typeTransaction);
  console.log("hash transaction: " + transactionHash);
  console.log("amountTransaction: " + amountTransaction);
  console.log("cryptoBuy: " + cryptoBuy);
  console.log("quantityBuy: " + quantityBuy);
  console.log("cryptoSell: " + cryptoSell);
  console.log("quantitySell: " + quantitySell);
  console.log("date: " + date);
  console.log("-------------------------------");
}
