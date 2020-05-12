let scriptProperties = PropertiesService.getScriptProperties()
let CACHE = CacheService.getScriptCache()

const OPENAPI_TOKEN = scriptProperties.getProperty('OPENAPI_TOKEN')
const TRADING_START_AT = new Date('Apr 01, 2020 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24

function isoToDate(dateStr){
  // How to format date string so that google scripts recognizes it?
  // https://stackoverflow.com/a/17253060
  let str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
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
    let url = this.baseUrl + methodUrl
    Logger.log(`[API Call] ${url}`)
    let params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
    let response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText())
  }
  
  getInstrumentByTicker(ticker) {
    let url = `market/search/by-ticker?ticker=${ticker}`
    let data = this._makeApiCall(url)
    return data.payload.instruments[0]
  }
  
  getOrderbookByFigi(figi) {
    let url = `market/orderbook?depth=1&figi=${figi}`
    let data = this._makeApiCall(url)
    return data.payload
  }
  
  getOperations(from, to, figi) {
    // Arguments `from` && `to` should be in ISO 8601 format
    let url = `operations?from=${from}&to=${to}&figi=${figi}`
    let data = this._makeApiCall(url)
    return data.payload.operations
  }
}

let tinkoffClient = new TinkoffClient(OPENAPI_TOKEN)

function _getFigiByTicker(ticker) {
  let cached = CACHE.get(ticker)
  if (cached != null) 
    return cached
  let instrument = tinkoffClient.getInstrumentByTicker(ticker)
  let figi = instrument.figi
  CACHE.put(ticker, figi)
  return figi
}

function getPriceByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  let figi = _getFigiByTicker(ticker)
  let orderbook = tinkoffClient.getOrderbookByFigi(figi)
  return orderbook.lastPrice
}

function _calculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    t = trades[j]
    totalQuantity += t.quantity
    totalSum += t.quantity * t.price
  }
  let weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}

function getTrades(ticker, from, to) {
  let figi = _getFigiByTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    let now = new Date()
    to = new Date(now + MILLIS_PER_DAY)
    to = to.toISOString()
  }
  let operations = tinkoffClient.getOperations(from, to, figi)
  
  let values = [
    ["ID", "Date", "Operation", "Ticker", "Quantity", "Price", "Currency", "SUM", "Commission"], 
  ]
  for (let i=operations.length-1; i>=0; i--) {
    let op = operations[i]
    if (op.operationType == "BrokerCommission" || op.status == "Decline") {continue}
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(op.trades) // calculate weighted values
    if (op.operationType == "Buy") {  // inverse values in a way, that it will be easier to work with
      totalQuantity = -totalQuantity
      totalSum = -totalSum
    }
    values.push([
      op.id, isoToDate(op.date), op.operationType, ticker, totalQuantity, weigthedPrice, op.currency, totalSum, op.commission.value
    ])
  }
  return values
}

function onEdit(e)
{
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  sheet.getRange('Z1').setValue(Math.random())
}
