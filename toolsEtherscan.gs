const apiScanInfos = {
  etherscan: {
    methods:{
      swapExactETHForTokens: {
          methodId: "0x7ff36ab5",
          logIndexCryptoBuy: 2,
          logIndexQuantityBuy: 2,
          crytoSell: "Ethereum",
          logIndexQuantitySell: 0,
      },
      swapExactTokensForETH: {
        methodId: "0x18cbafe5",
        cryptoBuy: "Ethereum",
        logIndexQuantityBuy: 1,
        logIndexCryptoSell: 0,
        logIndexQuantitySell: 0,
      },
      unibotBuyV2: {
        methodId: "0x19948479",
        logIndexCryptoBuy: 2,
        logIndexQuantityBuy: 3,
        crytoSell: "Ethereum",
        logIndexQuantitySell: 0,
      },
      unibotSellV2: {
        methodId: "0x8ee938a9",
        cryptoBuy: "Ethereum",
        logIndexQuantityBuy: 2,
        logIndexCryptoSell: 0,
        logIndexQuantitySell: [0, 1]
      },
      unibotSellV3: {
      methodId: "0x8107aee3",
      cryptoBuy: "Ethereum",
      logIndexQuantityBuy: 2,
      logIndexCryptoSell: 0,
      logIndexQuantitySell: [0, 1]
      }
    },
    nameDomain: "etherscan.io",
    apiKey: "5DDYRKH9WW2CASRHPBR2RZ29467S31V18D"
  }
}
/*
const apiScanInfos = {
  bscscan: {
    methodId: [
      { methodIdSwapExactETHForTokens: "0x7ff36ab5" },
      { methodIdMulticallBuy: "0x5ae401dc" },
      { methodIdSwapExactTokensForETH: "0x18cbafe5" },
      { methodIdSwapExactTokensForTokens: "0x38ed1739" },
      { methodIdSwapExactETHForTokensSupportingFeeOnTransferTokens: "0xb6f9de95" },
      { methodIdSwapExactTokensForTokensSupportingFeeOnTransferTokens: "0x5c11d795" }],
    nameDomain: "bscscan.com",
    apiKey: "T1ZB5E35GYX7GD9M72RT2CKNWPJ6574GND"
  }
}*/
function addTransactionsEtherscan() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let apiKeyEtherscan = "5DDYRKH9WW2CASRHPBR2RZ29467S31V18D"
  let addressses = ["0x897cbe1b142eA9b34A82A7302c84AD19F73A1C70"/*, "0x914d449A0d989F2CCF52975ced45F925828c8030"*/]
  let listSellTransactions;
  let listBuyTransactions;
  const cryptoNameInGS = getAllCryptoNameInGS();

  for (address of addressses) {
    listSellTransactions = [];
    listBuyTransactions = [];

    let allTransaction = fetchDataFromEtherscan(address, "etherscan.io");
    if (allTransaction) {
      if (allTransaction.listBuyTransactions.length == 0 && allTransaction.listSellTransactions.length == 0) {
        console.log("|main| listBuyTransactions.length == 0 && listSellTransactions.length == 0");
        return null
      }
      if (allTransaction.listBuyTransactions) {
        listBuyTransactions = allTransaction.listBuyTransactions;
      }
      if (allTransaction.listSellTransactions) {
        listSellTransactions = allTransaction.listSellTransactions;
      }
    } else {
      console.log("|main| Aucune Transaction");
      return null;
    }

    let listTransactionsToAdd = getListTransactionToAdd(listBuyTransactions, listSellTransactions);
    let listBuyTransactionsToAdd = listTransactionsToAdd.listBuyTransactionsToAdd;
    let listSellTransactionsToAdd = listTransactionsToAdd.listSellTransactionsToAdd;

    for (buyTransactionToAdd of listBuyTransactionsToAdd) {
      console.log("|main| Ajout de la transaction: " + buyTransactionToAdd.hash + " , dans la feuille google sheet");
      if (!(cryptoNameInGS.includes(buyTransactionToAdd.cryptoBuy))) {
        console.log("|main| Ajout de: " + buyTransactionToAdd.cryptoBuy + " à la feuille google sheet")
        //addNewCrypto(buyTransactionToAdd.cryptoBuy);
        cryptoNameInGS.push(buyTransactionToAdd.cryptoBuy);
      }
      if (!(cryptoNameInGS.includes(buyTransactionToAdd.cryptoSell))) {
        console.log("|main| Ajout de: " + buyTransactionToAdd.cryptoSell + " à la feuille google sheet")
        //addNewCrypto(buyTransactionToAdd.cryptoSell);
        cryptoNameInGS.push(buyTransactionToAdd.cryptoSell);
      }
      //transaction(buyTransactionToAdd.amountTransaction, buyTransactionToAdd.cryptoBuy, buyTransactionToAdd.quantityBuy, buyTransactionToAdd.cryptoSell, buyTransactionToAdd.quantitySell, buyTransactionToAdd.date)
    }
    for (sellTransactionToAdd of listSellTransactionsToAdd) {
      console.log("|main| Ajout de la transaction: " + sellTransactionToAdd.hash + " , dans la feuille google sheet");
      if (!(cryptoNameInGS.includes(sellTransactionToAdd.cryptoBuy))) {
        console.log("|main| Ajout de: " + sellTransactionToAdd.cryptoBuy + " à la feuille google sheet")
        //addNewCrypto(sellTransactionToAdd.cryptoBuy);
        cryptoNameInGS.push(sellTransactionToAdd.cryptoBuy);
      }
      if (!(cryptoNameInGS.includes(sellTransactionToAdd.cryptoSell))) {
        console.log("|main| Ajout de: " + sellTransactionToAdd.cryptoSell + " à la feuille google sheet")
        //addNewCrypto(sellTransactionToAdd.cryptoSell);
        cryptoNameInGS.push(sellTransactionToAdd.cryptoSell);
      }
      //transaction(sellTransactionToAdd.amountTransaction, sellTransactionToAdd.cryptoBuy, sellTransactionToAdd.quantityBuy, sellTransactionToAdd.cryptoSell, sellTransactionToAdd.quantitySell, sellTransactionToAdd.date)
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

  const methodIdSwapExactETHForTokens = "0x7ff36ab5";
  const methodIdSwapExactTokensForETH = "0x18cbafe5";
  const methodIdUnibotBuyV2 = "0x19948479";
  const methodIdUnibotSellV2 = "0x8ee938a9";
  const methodIdUnibotSellV3 = "0x8107aee3";

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const lastTimeStampTransaction = getLastTimeStampTransaction();
  let listBuyTransactions = [];
  let listSellTransactions = [];
  let timestampMax = 0;
  let blockNumberMax = 0;

  if (data.status === "1" && data.message === "OK") {

    const transactions = data.result;
    for (const txHash in transactions) {
      const transaction = transactions[txHash];
      if (transaction.timeStamp > timestampMax) {
        timestampMax = transaction.timeStamp;
      }
      if (transaction.blockNumber > blockNumberMax) {
        blockNumberMax = transaction.blockNumber;
      }
      if ((transaction.methodId == methodIdSwapExactETHForTokens || transaction.methodId == methodIdUnibotBuyV2) && transaction.timeStamp > lastTimeStampTransaction && transaction.txreceipt_status == 1) {
        listBuyTransactions.push(transaction);
      }
      else if ((transaction.methodId == methodIdSwapExactTokensForETH || transaction.methodId == methodIdUnibotSellV2 || transaction.methodId == methodIdUnibotSellV3) && transaction.timeStamp > lastTimeStampTransaction && transaction.txreceipt_status == 1) {
        listSellTransactions.push(transaction);
      }
    }
    console.log("|fetchDataFromEtherscan| listBuyTransactions.length: " + listBuyTransactions.length);
    console.log("|fetchDataFromEtherscan| listSellTransactions.length: " + listSellTransactions.length);
    //setLastTimeStampTransaction(timestampMax);

    return {
      listBuyTransactions: listBuyTransactions,
      listSellTransactions: listSellTransactions
    }
  } else {
    console.log("|fetchDataFromEtherscan| API Etherscan message: " + data.message);
    return null;
  }
}

function getListTransactionToAdd(listBuyTransactions, listSellTransactions) {
  let amountTransaction;
  let quantitySell;
  let quantityBuy;
  let cryptoSell;
  let cryptoBuy;
  let date;
  let logs;
  let listBuyTransactionsToAdd = [];
  let listSellTransactionsToAdd = [];

  const ethPrice = getCurrentEthPrice();

  for (buyTransaction of listBuyTransactions) {
    let buyTransactionReceipt = getTransactionReceipt(buyTransaction.hash);
    if (buyTransactionReceipt) {
      logs = buyTransactionReceipt.logs;
      if (logs) {
        if (logs.length == 5) {
          quantitySell = hexToDecimal(logs[0].data) / 10e17;
          quantityBuy = hexToDecimal(logs[2].data) / 10e17;
          amountTransaction = ethPrice * (quantitySell + (buyTransaction.gasUsed * (buyTransaction.gasPrice / 10e17)));
          console.log("amountTransaction: "+amountTransaction+" | hash: "+buyTransaction.hash)
          cryptoSell = "Ethereum";
          cryptoBuy = getCryptpoName(logs[2].address);
          date = getFormattedDate(buyTransaction.timeStamp);
          listBuyTransactionsToAdd.push({ hash: buyTransaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          //showTransactionDetails("ACHAT 5 LOGS", logs[0].transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);

        } else if (logs.length == 6) {
          quantitySell = hexToDecimal(logs[0].data) / 10e17;
          quantityBuy = hexToDecimal(logs[3].data) / 10e17;
          amountTransaction = ethPrice * (quantitySell + (buyTransaction.gasUsed * (buyTransaction.gasPrice / 10e17)));
          cryptoSell = "Ethereum";
          cryptoBuy = getCryptpoName(logs[2].address);
          date = getFormattedDate(buyTransaction.timeStamp);
          listBuyTransactionsToAdd.push({ hash: buyTransaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          //showTransactionDetails("ACHAT 6 LOGS", logs[0].transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);

        } else {
          console.log("|getListTransactionToAdd| Achat, logs length != 5 & 6");
        }
      } else {
        console.log("|getListTransactionToAdd| Achat, logs :" + buyTransaction.hash + "is null")
      }
    } else {
      console.log("|getListTransactionToAdd| Achat, buyTransactionReceipt of: " + buyTransaction.hash + "is null")
    }
  }

  for (sellTransaction of listSellTransactions) {
    let sellTransactionReceipt = getTransactionReceipt(sellTransaction.hash);
    if (sellTransactionReceipt) {
      logs = sellTransactionReceipt.logs;
      if (logs) {
        if (logs.length == 5) {
          quantityBuy = hexToDecimal(logs[1].data) / 10e17;
          amountTransaction = ethPrice * (quantityBuy + (sellTransaction.gasUsed * (sellTransaction.gasPrice / 10e17)));
          cryptoBuy = "Ethereum";
          cryptoSell = getCryptpoName(logs[0].address);
          if (cryptoSell == "Tether USDt") {
            quantitySell = hexToDecimal(logs[0].data) / 10e5;
          } else {
            quantitySell = hexToDecimal(logs[0].data) / 10e17;
          }

          date = getFormattedDate(sellTransaction.timeStamp);

          listSellTransactionsToAdd.push({ hash: sellTransaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          //showTransactionDetails("VENTE 5 LOGS", logs[0].transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);
        } else if (logs.length == 6) {
          quantitySell =
            hexToDecimal(logs[0].data) / 10e17 +
            hexToDecimal(logs[1].data) / 10e17;
          quantityBuy = hexToDecimal(logs[2].data) / 10e17;
          amountTransaction = ethPrice * (quantityBuy + (sellTransaction.gasUsed * (sellTransaction.gasPrice / 10e17)));
          cryptoBuy = "Ethereum";
          cryptoSell = getCryptpoName(logs[0].address);
          date = getFormattedDate(sellTransaction.timeStamp);
          listSellTransactionsToAdd.push({ hash: sellTransaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
          //showTransactionDetails("VENTE 6 LOGS", logs[0].transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);

        } else {
          console.log("|getListTransactionToAdd| Vente, logs length != 5 & 6");
        }
      } else {
        console.log("|getListTransactionToAdd| Vente, logs :" + sellTransaction.hash + "is null")
      }
    } else {
      console.log("|getListTransactionToAdd| Vente, sellTransactionReceipt of: " + sellTransaction.hash + "is null")
    }
  }
  return {
    listBuyTransactionsToAdd: listBuyTransactionsToAdd,
    listSellTransactionsToAdd: listSellTransactionsToAdd
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
      return name;
    } else {
      console.log("|getCryptpoName| name is null dernier else");
    }
  }
  return null;
}

function getCurrentEthPrice() {
  const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?id=1027";
  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": "9c8a01c2-519b-4173-8c5b-4842461e61e4",
    },
    muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
  };
  let quoteResponse = UrlFetchApp.fetch(urlQuotes, options);

  if (quoteResponse.getResponseCode() === 200) {
    let quoteJsonData = quoteResponse.getContentText();
    const quoteJson = JSON.parse(quoteJsonData);
    const cryptoPrice = quoteJson.data["1027"].quote.USD.price;
    return cryptoPrice;
  }
  else {
    console.log("|getCurrentEthPrice| Fail request cmc api,quoteResponse.getResponseCode(): " + quoteResponse.getResponseCode())
    return null;
  }

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
  console.log("VtypeTransaction: " + typeTransaction);
  console.log("hash transaction: " + transactionHash);
  console.log("amountTransaction: " + amountTransaction);
  console.log("cryptoBuy: " + cryptoBuy);
  console.log("quantityBuy: " + quantityBuy);
  console.log("cryptoSell: " + cryptoSell);
  console.log("quantitySell: " + quantitySell);
  console.log("date: " + date);
  console.log("-------------------------------");
}
