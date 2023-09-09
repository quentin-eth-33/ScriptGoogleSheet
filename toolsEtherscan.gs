/*
-appel api cmc pour avoir le nom de la crypto ne fonctionne pas
-vérifier l193
*/


// Variables Globales:
const scriptProperties = PropertiesService.getScriptProperties();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const apiKeyEtherscan = scriptProperties.getProperty("apiKeyEtherscan");
const apiKeyBscscan = scriptProperties.getProperty("apiKeyBscscan");
const apiKeyCMC = scriptProperties.getProperty("apiKeyCMC");
//https://api.bscscan.com/api?module=account&action=txlist&address=0x897cbe1b142eA9b34A82A7302c84AD19F73A1C70&startblock=0&endblock=99999999&sort=desc&apikey=T1ZB5E35GYX7GD9M72RT2CKNWPJ6574GND
function listenerBlockchain() {
  const addressses = ["0x897cbe1b142eA9b34A82A7302c84AD19F73A1C70"/*, "0x914d449A0d989F2CCF52975ced45F925828c8030"*/]
  const cryptoNameInGS = getAllCryptoNameInGS();

  const etherscanInfos = {
    apiKey: apiKeyEtherscan,
    domainName: "etherscan.io",
    lastTimestamp: parseInt(scriptProperties.getProperty("lastTimestampEtherscan")),
    lastBlock: parseInt(scriptProperties.getProperty("lastBlockEtherscan")),
    cryptoFee: "Ethereum"
  }

  const bscscanInfos = {
    apiKey: apiKeyBscscan,
    domainName: "bscscan.com",
    lastTimestamp: parseInt(scriptProperties.getProperty("lastTimestampBscscan")),
    lastBlock: parseInt(scriptProperties.getProperty("lastBlockBscscan")),
    cryptoFee: "BNB"
  }

  const blockchains = [/*etherscanInfos,*/ bscscanInfos];

  for (const address of addressses) {
    for (const blockchain of blockchains) {
      let listTransactionsToAdd = fetchDataFromEtherscan(address, blockchain);
      if (listTransactionsToAdd) {
        if (listTransactionsToAdd.length > 0) {
          console.log("allTransaction.length: " + listTransactionsToAdd.length);
          for (transactionToAdd of listTransactionsToAdd) {
            console.log("|listenerBlockchain| Ajout de la transaction: " + transactionToAdd.hash + " , dans la feuille google sheet");
            if (!(cryptoNameInGS.includes(transactionToAdd.cryptoBuy))) {
              console.log("|listenerBlockchain| Ajout de: " + transactionToAdd.cryptoBuy + " à la feuille google sheet")
              //addNewCrypto(transactionToAdd.cryptoBuy);
              cryptoNameInGS.push(transactionToAdd.cryptoBuy);
            }
            if (!(cryptoNameInGS.includes(transactionToAdd.cryptoSell))) {
              console.log("|listenerBlockchain| Ajout de: " + transactionToAdd.cryptoSell + " à la feuille google sheet")
              //addNewCrypto(transactionToAdd.cryptoSell);
              cryptoNameInGS.push(transactionToAdd.cryptoSell);
            }
            //transaction(transactionToAdd.amountTransaction, transactionToAdd.cryptoBuy, transactionToAdd.quantityBuy, transactionToAdd.cryptoSell, transactionToAdd.quantitySell, transactionToAdd.date)
          }
        } else {
          console.log("|listenerBlockchain| Pas de nouvelles transactions à ajouter")
        }
      }
      else {
        console.log("|listenerBlockchain| Pas de nouvelles transactions à ajouter")
      }
    }

  }
}
//https://api.bscscan.com/api?module=account&action=txlist&address=0x897cbe1b142eA9b34A82A7302c84AD19F73A1C70&startblock=0&endblock=99999999&sort=desc&apikey=T1ZB5E35GYX7GD9M72RT2CKNWPJ6574GND
function fetchDataFromEtherscan(address, blockchain) {

  const transferTopic = "0xddf252ad1be2c89b69c2b068fc378daa952ba7f163c4a11628f55a4df523b3ef";
  const depositTopic = "0xe1fffcc4923d04b559f4d29a8bfc6cda04eb5b0d3c460751c2402c5c5cc9109c";
  const withdrawTopic = "0x7fcf532c15f0a6db0bd6d0e038bea71d30d808c7d98cb3bf7268a95bf5081b65";
  const methodIdSwapExactETHForTokens = "0x7ff36ab5";
  const methodIdSwapExactTokensForETH = "0x18cbafe5";
  const methodIdUnibotBuyV2 = "0x19948479";
  const methodIdUnibotSellV2 = "0x8ee938a9";
  const methodIdUnibotSellV3 = "0x8107aee3";
  const methodIdswapExactETHForTokensSupportingFeeOnTransferTokens = "0xb6f9de95";
  const methodIdMulticall = "0x5ae401dc";
  const methodIdSwapExactTokensForTokens = "0x38ed1739";
  const methodIdSwapExactTokensForETHSupportingFeeOnTransferTokens = "0x791ac947";
  const methodIdExecute = "0x3593564c"

  const url = "https://api." + blockchain.domainName + "/api?module=account&action=txlist&address=" + address + "&startblock=" + blockchain.lastBlock + "&endblock=99999999&sort=desc&apikey=" + blockchain.apiKey;
  const options = {
    method: "get",
    muteHttpExceptions: true
  };

  const cryptoFeePrice = getCryptoPriceAndMC(blockchain.cryptoFee).cryptoPrice;

  const listMethodIdSwap = [methodIdSwapExactETHForTokens, methodIdSwapExactTokensForETH, methodIdUnibotBuyV2, methodIdUnibotSellV2, methodIdUnibotSellV3, methodIdswapExactETHForTokensSupportingFeeOnTransferTokens, methodIdMulticall, methodIdSwapExactTokensForTokens, methodIdSwapExactTokensForETHSupportingFeeOnTransferTokens, methodIdExecute]

  const lastTimeStampTransaction = blockchain.lastTimestamp;
  let timestampMax = 0;
  let blockNumberMax = 0;
  let transactionReceipt;
  let amountTransaction;
  let quantitySell;
  let quantityBuy;
  let cryptoSell;
  let cryptoBuy;
  let date;
  let logs;
  let gasFee;
  let listTransactionsToAdd = [];
  let listTransferOutLog = [];
  let listTransferInLog = [];
  let listDepositLog = [];
  let listWithdrawLog = [];
  let infosCryptoSell;
  let infosCryptoBuy;
  let addressTransferOut = address;
  let addressTransferIn = address;
  let addressCryptoOut;
  let addressCryptoIn;

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data.status === "1" && data.message === "OK") {

    const transactions = data.result;
    for (const transaction of transactions) {

      addressTransferIn = address;
      addressTransferOut = address;
      quantitySell = 0;
      quantityBuy = 0;
      listTransferOutLog = [];
      listTransferInLog = [];
      listDepositLog = [];
      listWithdrawLog = [];

      if (transaction.timeStamp > timestampMax) {
        timestampMax = transaction.timeStamp;
      }
      if (transaction.blockNumber > blockNumberMax) {
        blockNumberMax = transaction.blockNumber;
      }

      // A CHANGER LE 0
      if (listMethodIdSwap.includes(transaction.methodId) && transaction.timeStamp > /*blockchain.lastTimestamp*/1680193092 && transaction.txreceipt_status == 1) {

        transactionReceipt = getTransactionReceipt(transaction.hash, blockchain);
        if (transactionReceipt) {
          logs = transactionReceipt.logs;
          date = getFormattedDate(transaction.timeStamp);
          gasFee = transaction.gasUsed * (transaction.gasPrice / 10e17)

          for (const log of logs) {
            if (log.topics[0] == withdrawTopic) {
              listWithdrawLog.push(log);
            } else if (log.topics[0] == depositTopic) {
              listDepositLog.push(log);
            }
          }
          if (listWithdrawLog.length > 1 || listDepositLog.length > 1) {
            console.log("|fetchDataFromEtherscan| listWithdrawLog.length > 1 || listDepositLog.length > 1 | cas non traité | " + transaction.hash);
            continue;
          }
          else if (listDepositLog.length == 1) {
            addressTransferOut = getAddressValid(listDepositLog[0].topics[1]);
          }
          else if (listWithdrawLog.length == 1) {
            addressTransferIn = getAddressValid(listWithdrawLog[0].topics[1]);
          }

          for (const log of logs) {
            if (log.topics[0] == transferTopic) {
              if (getAddressValid(log.topics[1]).toLowerCase() == addressTransferOut.toLowerCase() || getAddressValid(log.topics[1]).toLowerCase() == address.toLowerCase()) {
                listTransferOutLog.push(log);
              }
              else if (getAddressValid(log.topics[2]).toLowerCase() == addressTransferIn.toLowerCase() || getAddressValid(log.topics[2]).toLowerCase() == address.toLowerCase()) {
                listTransferInLog.push(log);
              }
            }
          }

          if (listTransferOutLog.length > 0 && listTransferInLog.length > 0) {

            addressCryptoOut = listTransferOutLog[0].address;
            for (let i = 1; i < listTransferOutLog.length; i++) {
              if (!(listTransferOutLog[i].address == addressCryptoOut)) {
                console.log("|fetchDataFromEtherscan| listWithdrawLog.length > 1 || listDepositLog.length > 1 | cas non traité (plusieurs cryptos diff) | hash: " + transaction.hash);
                continue;
              }
            }

            cryptoSell = getCryptoName(listTransferOutLog[0].address);
            if (cryptoSell) {
              infosCryptoSell = getCryptoPriceAndMC(cryptoSell); // {cryptoPrice: 10 ,cryptoMC: 2};
              for (const transferOut of listTransferOutLog) {
                quantitySell += hexToDecimal(transferOut.data);
              }
              // Respectivement tether et usdc 
              if (cryptoSell == "0xdac17f958d2ee523a2206206994597c13d831ec7" || cryptoSell == "0xa0b86991c6218b36c1d19d4a2e9eb0ce3606eb48") {
                quantitySell = quantitySell / 10e6; // pas sur
              }
              else {
                quantitySell = quantitySell / 10e17;
              }
            }
            else {
              console.log("|fetchDataFromEtherscan| cryptoSell == null | hash: " + transaction.hash);
              continue;
            }

            addressCryptoIn = listTransferInLog[0].address;
            for (i = 1; i < listTransferInLog.length; i++) {
              if (!(listTransferInLog[i].address == addressCryptoIn)) {
                if (listTransferInLog[i].topics[1] != "0x0000000000000000000000000000000000000000000000000000000000000000") {
                  console.log("|fetchDataFromEtherscan| transfer in de plsr crypto diff, cas pas encore traité | hash: " + transaction.hash);
                  continue;
                }
              }
            }

            cryptoBuy = getCryptoName(listTransferInLog[0].address);
            if (cryptoBuy) {
              infosCryptoBuy = getCryptoPriceAndMC(cryptoBuy); // {cryptoPrice: 10 ,cryptoMC: 1};
              for (const transferIn of listTransferInLog) {
                if (transferIn.topics[1] != "0x0000000000000000000000000000000000000000000000000000000000000000") {
                  quantityBuy += hexToDecimal(transferIn.data);
                }
              }
              // Respectivement tether et usdc 
              if (cryptoBuy == "0xdac17f958d2ee523a2206206994597c13d831ec7" || cryptoBuy == "0xa0b86991c6218b36c1d19d4a2e9eb0ce3606eb48") {
                quantityBuy = quantityBuy / 10e6; // pas sur
              }
              else {
                quantityBuy = quantityBuy / 10e17;
              }
            }
            else {
              console.log("|fetchDataFromEtherscan| cryptoBuy == null | hash: " + transaction.hash);
              continue;
            }

            if (quantityBuy == 0 || quantityBuy == 0 || !cryptoSell || !cryptoBuy || (!infosCryptoBuy && !infosCryptoSell)) {
              console.log("|fetchDataFromEtherscan| quantityBuy == 0 ||quantityBuy == 0 || !cryptoSell || !cryptoBuy || !infosCryptoBuy ||!infosCryptoSell | hash: " + transaction.hash);
              continue;
            }

            amountTransaction = getAmountOfTransaction(infosCryptoBuy, infosCryptoSell, quantityBuy, quantitySell);
            if (amountTransaction) {
              amountTransaction += gasFee * cryptoFeePrice;
            }
            else {
              console.log("|fetchDataFromEtherscan| amountTransaction == null | hash: " + transaction.hash);
              continue;
            }

            if (cryptoSell && quantitySell && cryptoBuy && quantityBuy && transaction.hash && amountTransaction && date) {
              listTransactionsToAdd.push({ hash: transaction.hash, amountTransaction: amountTransaction.toString(), cryptoBuy: cryptoBuy, quantityBuy: quantityBuy.toString(), cryptoSell: cryptoSell, quantitySell: quantitySell.toString(), date: date })
            }
            else {
              console.log("|fetchDataFromEtherscan| Une des valeurs de la transaction est undefine | hash: " + transaction.hash);
              continue;
            }
            showTransactionDetails(transaction.hash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date);
          }
          else {
            console.log("|fetchDataFromEtherscan| listTransferOutLog.length > 0 && listWithdrawLog.length >0 non validé | hash: " + transaction.hash);
            continue;
          }
        }
      } else {
        console.log("|fetchDataFromEtherscan| transactionReceipt == null | hash" + transaction.hash);
        continue;
      }
    }
    //setLastTimeStampTransaction(timestampMax);
    return listTransactionsToAdd;
  } else {
    console.log("|fetchDataFromEtherscan| Fail call API Etherscan, message: " + data.message);
    return null;
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

function getTransactionReceipt(hash, blockchain) {
  let url = "https://api." + blockchain.domainName + "/api?module=proxy&action=eth_getTransactionReceipt&txhash=" + hash + "&apikey=" + blockchain.apiKey;
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

function getAddressValid(addressHex) {
  let address = "0x"
  address += addressHex.slice(26);
  return address.toString();
}

function getCryptoName(address) {
  let name;
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
    const url = "https://pro-api.coinmarketcap.com/v2/cryptocurrency/info?address=" + address;
    let options = {
      method: "get",
      headers: {
        "X-CMC_PRO_API_KEY": apiKeyCMC,
      },
      muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
    };
    let infoResponse = UrlFetchApp.fetch(url, options);

    if (infoResponse.getResponseCode() === 200) {
      let infoJsonData = JSON.parse(infoResponse.getContentText());
      let firstKey = Object.keys(infoJsonData.data)[0];
      name = infoJsonData.data[firstKey].name;
      if (name) {
        if (name == "WETH") {
          name = "Ethereum";
        } else if (name == "Wrapped BNB") {
          name = "BNB";
        }
        mapAddressCryptoGlobalString += `;${address}:${name}`;
        scriptProperties.setProperty('mapAddressCryptoGlobal', mapAddressCryptoGlobalString);
      } else {
        name = null;
        console.log("|getCryptoName| name is null premier else | address: " + address);
      }
    } else {
      name = null;
      console.log("|getCryptoName| Fail request cmc api, infoResponse.getResponseCode(): " + infoResponse.getResponseCode() + " | address: " + address);
      console.log("|getCryptoName| infoResponse.getContentText(): " + infoResponse.getContentText())
    }
  }
  else {
    name = mapAddressCryptoGlobal[address];
  }
  return name;
}

function getCryptoPriceAndMC(name) {
  const urlSymbols = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/map";
  const urlQuotes = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest";
  let cryptoId;
  let options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": apiKeyCMC,
    },
    muteHttpExceptions: true // Prevents throwing an exception for non-2xx responses
  };

  let cryptoPrice;
  let cryptoMC;
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
        console.log("|getCryptoPriceAndMC| Cryptomonnaie " + name + " introuvable");
        return null;
      }

      cryptoId = cryptoData.id;
      mapIdCryptoGlobalString += ";" + name + ":" + cryptoId;
      scriptProperties.setProperty('mapIdCryptoGlobal', mapIdCryptoGlobalString);
    } else if (response.getResponseCode() === 429) {
      console.log("|getCryptoPriceAndMC| nb limite de requete atteint | nom crypto: " + name)
      return null;
    } else {
      console.log("|getCryptoPriceAndMC| Erreur appel api cmc, response.getResponseCode(): " + response.getResponseCode() + " | nom crypto: " + name)
      return null;
    }
  }
  else {
    cryptoId = mapIdCryptoGlobal[name];
    if (!cryptoId) {
      console.log("|getCryptoPriceAndMC| cryptoId null | nom crypto: " + name)
      return null;
    }
  }

  let quoteResponse = UrlFetchApp.fetch(`${urlQuotes}?id=${cryptoId}`, options);
  if (quoteResponse.getResponseCode() === 200) {
    let quoteJsonData = quoteResponse.getContentText();
    const quoteJson = JSON.parse(quoteJsonData);

    cryptoPrice = quoteJson.data[cryptoId].quote.USD.price;
    cryptoMC = quoteJson.data[cryptoId].quote.USD.market_cap;
    if (!cryptoPrice || !cryptoMC) {
      console.log("|getCryptoPriceAndMC| !cryptoPrice || !cryptoMC non validé | nom crypto: " + name)
      return null;
    }
  } else if (quoteResponse.getResponseCode() === 429) {
    console.log("|getCryptoPriceAndMC| nb limite de requete atteint | nom crypto: " + name)
    return null;
  } else {
    console.log("|getCryptoPriceAndMC| Erreur appel api cmc, quoteResponse.getResponseCode(): " + quoteResponse.getResponseCode() + " | nom crypto: " + name)
    return null;
  }

  return {
    cryptoPrice: cryptoPrice,
    cryptoMC: cryptoMC
  }
}

function getAllCryptoNameInGS() {
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

function getAmountOfTransaction(cryptoInfosBuy, cryptoInfosSell, quantityBuy, quantitySell) {
  let amount;
  if (cryptoInfosBuy) {
    if (cryptoInfosSell) {
      if (cryptoInfosBuy.cryptoMC > cryptoInfosSell.cryptoMC) {
        amount = cryptoInfosBuy.cryptoPrice * quantityBuy;
      } else {
        amount = cryptoInfosSell.cryptoPrice * quantitySell;
      }
    } else {
      amount = cryptoInfosBuy.cryptoPrice * quantityBuy;
    }
  }
  else if (cryptoInfosSell) {
    amount = cryptoInfosSell.cryptoPrice * quantitySell;
  }
  else {
    amount = null;
  }
  return amount;
}

function showTransactionDetails(transactionHash, amountTransaction, cryptoBuy, quantityBuy, cryptoSell, quantitySell, date) {
  console.log("hash transaction: " + transactionHash);
  console.log("amountTransaction: " + amountTransaction);
  console.log("cryptoBuy: " + cryptoBuy);
  console.log("quantityBuy: " + quantityBuy);
  console.log("cryptoSell: " + cryptoSell);
  console.log("quantitySell: " + quantitySell);
  console.log("date: " + date);
  console.log("-------------------------------");
}
