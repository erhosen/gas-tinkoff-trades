var scriptProperties = PropertiesService.getScriptProperties()
var CACHE = CacheService.getScriptCache()

const OPENAPI_TOKEN = scriptProperties.getProperty('OPENAPI_TOKEN')
const TRADING_START_AT = new Date('Apr 01, 2020 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24

function isoToDate(dateStr){
  // How to format date string so that google scripts recognizes it?
  // https://stackoverflow.com/a/17253060
  var str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
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
    var url = this.baseUrl + methodUrl
    Logger.log(`[API Call] ${url}`)
    var params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
    var response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText())
  }
  
  getInstrumentByTicker(ticker) {
    var url = `market/search/by-ticker?ticker=${ticker}`
    var data = this._makeApiCall(url)
    return data.payload.instruments[0]
  }
  
  getOrderbookByFigi(figi) {
    var url = `market/orderbook?depth=1&figi=${figi}`
    var data = this._makeApiCall(url)
    return data.payload
  }
  
  getOperations(from, to, figi) {
    // Arguments `from` && `to` should be in ISO 8601 format
    var url = `operations?from=${from}&to=${to}&figi=${figi}`
    var data = this._makeApiCall(url)
    return data.payload.operations
  }
}

var tinkoffClient = new TinkoffClient(OPENAPI_TOKEN)

function _getFigiByTicker(ticker) {
  var cached = this.CACHE.get(ticker)
  if (cached != null) 
    return cached
  var instrument = tinkoffClient.getInstrumentByTicker(ticker)
  var figi = instrument.figi
  CACHE.put(ticker, figi)
  return figi
}

function getPriceByTicker(ticker, dummy) {
  // dummy attribute uses for auto-refreshing the value each time the sheet is updating.
  // see https://stackoverflow.com/a/27656313
  var figi = _getFigiByTicker(ticker)
  var orderbook = tinkoffClient.getOrderbookByFigi(figi)
  return orderbook.lastPrice
}

function getTrades(ticker, from, to) {
  var figi = _getFigiByTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    var now = new Date()
    to = new Date(now + MILLIS_PER_DAY)
    to = to.toISOString()
  }
  var operations = tinkoffClient.getOperations(from, to, figi)
  
  var values = [
    ["ID", "Date", "Operation", "Ticker", "Quantity", "Price", "Currency", "SUM", "Commission"], 
  ]
  for (var i=operations.length-1; i>=0; i--) {
    var op = operations[i]
    if (op.operationType == "BrokerCommission" || op.status == "Decline") {continue}
    var wSum = 0
    var totalQuantity = 0
    for (var j in op.trades) {
      t = op.trades[j]
      totalQuantity += t.quantity
      wSum += t.quantity * t.price
    }
    var weigthedAvg = wSum / totalQuantity
    if (op.operationType == "Buy") {
      totalQuantity = -totalQuantity
      wSum = -wSum
    }
    var commission = op.commission.value
    values.push([op.id, isoToDate(op.date), op.operationType, ticker, totalQuantity, weigthedAvg, op.currency, wSum, commission])
  }
  return values
}

function onEdit(e)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  sheet.getRange('Z1').setValue(Math.random())
}
