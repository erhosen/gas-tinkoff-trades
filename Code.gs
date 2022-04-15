/** @OnlyCurrentDoc */

const scriptProperties = PropertiesService.getScriptProperties()
const OPENAPI_TOKEN = scriptProperties.getProperty('OPENAPI_TOKEN')

const CACHE = CacheService.getScriptCache()
const CACHE_MAX_AGE = 21600 // 6 Hours

const TRADING_START_AT = new Date('Apr 01, 2020 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24

function isoToDate(dateStr){
  // How to format date string so that google scripts recognizes it?
  // https://stackoverflow.com/a/17253060
  const str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
  return new Date(str)
}

class TinkoffClient {
  // Doc: https://tinkoffcreditsystems.github.io/invest-openapi/swagger-ui/
  // How to create a token: https://tinkoffcreditsystems.github.io/invest-openapi/auth/
  constructor(token) {
    this.token = token
    this.baseUrl = 'https://api-invest.tinkoff.ru/openapi/'
  }
  
  _makeApiCall(methodUrl) {
    const url = this.baseUrl + methodUrl
    Logger.log(`[API Call] ${url}`)
    const params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
    const response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText())
  }
  
  getInstrumentByTicker(ticker) {
    const url = `market/search/by-ticker?ticker=${ticker}`
    const data = this._makeApiCall(url)
    return data.payload
  }
  
  getOrderbookByFigi(figi, depth) {
    const url = `market/orderbook?depth=${depth}&figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload
  }
  
  getOperations(from, to, figi) {
    // Arguments `from` && `to` should be in ISO 8601 format
    const url = `operations?from=${from}&to=${to}&figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload.operations
  }
}

const tinkoffClient = new TinkoffClient(OPENAPI_TOKEN)

function _getFigiByTicker(ticker) {
  const CACHE_KEY_PREFIX = 'figi_'
  const ticker_cache_key = CACHE_KEY_PREFIX + ticker

  const cached = CACHE.get(ticker)
  if (cached != null) 
    return cached
  const {instruments,total} = tinkoffClient.getInstrumentByTicker(ticker)
  if (total > 0) {
    const figi = instruments[0].figi
    CACHE.put(ticker_cache_key, figi, CACHE_MAX_AGE)
    return figi
  } else {
    return null
  }
}

/**
 * Получение последней цены инструмента по тикеру
 * @param {"GAZP"} ticker Тикер инструмента
 * @return {}             Last price
 * @customfunction
 */
function getPriceByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {lastPrice} = tinkoffClient.getOrderbookByFigi(figi, 1)
  return lastPrice
}

/**
 * Получение Bid/Ask спреда инструмента по тикеру
 * @param {"GAZP"} ticker Тикер инструмента
 * @return {0.03}         Спред в %
 * @customfunction
 */
function getBidAskSpreadByTicker(ticker) { // dummy parameter is optional
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,bids,asks} = tinkoffClient.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return (asks[0].price-bids[0].price) / asks[0].price
  else
    return null
}

function getMaxBidByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,bids} = tinkoffClient.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return [
    ["Max bid", "Quantity"],
    [bids[0].price, bids[0].quantity]
    ]
  else
    return null
}

function getMinAskByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  const figi = _getFigiByTicker(ticker)
  const {tradeStatus,asks} = tinkoffClient.getOrderbookByFigi(figi, 1)
  if (tradeStatus != 'NotAvailableForTrading')
    return [
      ["Min ask", "Quantity"],
      [asks[0].price, asks[0].quantity]
    ]
  else
    return null
}

function _calculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    const {quantity, price} = trades[j]
    totalQuantity += quantity
    totalSum += quantity * price
  }
  const weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}

/**
 * Получение списка операций по тикеру инструмента
 * @param {String} ticker Тикер инструмента для фильтрации
 * @param {String} from Начальная дата
 * @param {String} to Конечная дата
 * @return {Array} Массив результата
 * @customfunction
 */
function getTrades(ticker, from, to) {
  const figi = _getFigiByTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    const now = new Date()
    to = new Date(now + MILLIS_PER_DAY)
    to = to.toISOString()
  }
  const operations = tinkoffClient.getOperations(from, to, figi)
  
  const values = [
    ["ID", "Date", "Operation", "Ticker", "Quantity", "Price", "Currency", "SUM", "Commission"], 
  ]
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, id, date, currency, commission} = operations[i]
    if (operationType == "BrokerCommission" || status == "Decline" || operationType == "Dividend")
      continue
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades) // calculate weighted values
    if (operationType == "Buy") {  // inverse values in a way, that it will be easier to work with
      totalQuantity = -totalQuantity
      totalSum = -totalSum
    }
    let com_val = 0
    if (commission){
      com_val = commission.value
    }else{
      com_val = null
    }
    values.push([
      id, isoToDate(date), operationType, ticker, totalQuantity, weigthedPrice, currency, totalSum, com_val
    ])
  }
  return values
}

/**
 * Добавляет меню с командой вызова функции обновления значений служебной ячейки (для обновления вычислнений функций, ссылающихся на эту ячейку)
 *
 **/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var entries = [{
    name : "Обновить",
    functionName : "refresh"
  }]
  sheet.addMenu("TI", entries)
};

function refresh() {
  SpreadsheetApp.getActiveSpreadsheet().getRange('Z1').setValue(new Date());
}
