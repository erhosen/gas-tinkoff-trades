/** @OnlyCurrentDoc */

const scriptProperties = PropertiesService.getScriptProperties()
const OPENAPI_TOKEN = scriptProperties.getProperty('OPENAPI_TOKEN')

const CACHE = CacheService.getScriptCache()
const CACHE_MAX_AGE = 21600 // 6 Hours

const TRADING_START_AT = new Date('Apr 01, 2020 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24

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

  getPortfolio(){
    const url = `portfolio`
    const data = this._makeApiCall(url)
    return data.payload.positions
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
 * Получение портфеля
 * @return {Array}     Массив с результатами
 * @customfunction
 */
function getPortfolio() {
  const portfolio = tinkoffClient.getPortfolio()
  const values = []
  values.push(["Тикер","Название","Тип","Кол-во","Ср.цена покупки","Ст-ть покупки","Доходность","Тек.ст-ть","НКД","Валюта"])
  for (let i=0; i<portfolio.length; i++) {
    let {ticker, name, instrumentType, balance, averagePositionPrice, averagePositionPriceNoNkd, expectedYield} = portfolio [i]
    let NKD=null
    if (averagePositionPriceNoNkd){
      NKD = averagePositionPrice.value - averagePositionPriceNoNkd.value
      averagePositionPrice.value = averagePositionPriceNoNkd.value
    }

    values.push([
      ticker, name, instrumentType, balance, averagePositionPrice.value, averagePositionPrice.value * balance, expectedYield.value, averagePositionPrice.value * balance + expectedYield.value, NKD, averagePositionPrice.currency
    ])
  }
  return values
}

/**  ============================== Tinkoff V2 ==============================
*
* https://tinkoff.github.io/investAPI/
*
**/
class _TinkoffClientV2 {
  constructor(token){
    this.token = token
    this.baseUrl = 'https://invest-public-api.tinkoff.ru/rest/'
    //Logger.log(`[_TinkoffClientV2.constructor]`)
  }
  _makeApiCall(methodUrl,data){
    const url = this.baseUrl + methodUrl
    Logger.log(`[Tinkoff OpenAPI V2 Call] ${url}`)
    const params = {
      'method': 'post',
      'headers': {'accept': 'application/json', 'Authorization': `Bearer ${this.token}`},
      'contentType': 'application/json',
      'payload' : JSON.stringify(data)}
    
    const response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText())
  }
  // ----------------------------- InstrumentsService -----------------------------
  _Bonds(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Bonds`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Shares(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Shares`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Futures(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Futures`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Etfs(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Etfs`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _Currencies(instrumentStatus) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/Currencies`
    const data = this._makeApiCall(url, {'instrumentStatus': instrumentStatus})
    return data
  }
  _GetInstrumentBy(idType,classCode,id) {
    const url = `tinkoff.public.invest.api.contract.v1.InstrumentsService/GetInstrumentBy`
    const data = this._makeApiCall(url, {'idType': idType, 'classCode': classCode, 'id': id})
    return data
  }
  // ----------------------------- MarketDataService -----------------------------
  _GetLastPrices(figi_arr) {
    const url = 'tinkoff.public.invest.api.contract.v1.MarketDataService/GetLastPrices'
    const data = this._makeApiCall(url,{'figi': figi_arr})
    return data
  }
  _GetOrderBookByFigi(figi,depth) {
    const url = `tinkoff.public.invest.api.contract.v1.MarketDataService/GetOrderBook`
    const data = this._makeApiCall(url,{'figi': figi, 'depth': depth})
    return data
  }
  // ----------------------------- OperationsService -----------------------------
  _GetOperations(accountId,from,to,state,figi) {
    const url = 'tinkoff.public.invest.api.contract.v1.OperationsService/GetOperations'
    const data = this._makeApiCall(url,{'accountId': accountId,'from': from,'to': to,'state': state,'figi': figi})
    return data
  }
  _GetPortfolio(accountId) {
    const url = 'tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio'
    const data = this._makeApiCall(url,{'accountId': accountId})
    return data
  }
  // ----------------------------- UsersService -----------------------------
  _GetAccounts() {
    const url = 'tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts'
    const data = this._makeApiCall(url,{})
    return data
  }
  _GetInfo() {
    const url = 'tinkoff.public.invest.api.contract.v1.UsersService/GetInfo'
    const data = this._makeApiCall(url,{})
    return data
  }
}

const tinkoffClientV2 = new _TinkoffClientV2(OPENAPI_TOKEN)

function _GetTickerNameByFIGI(figi) {
  //Logger.log(`[TI_GetTickerByFIGI] figi=${figi}`)   // DEBUG
  const {ticker,name} = tinkoffClientV2._GetInstrumentBy('INSTRUMENT_ID_TYPE_FIGI',null,figi).instrument
  return [ticker,name]
}

function TI_GetLastPrice(ticker) {
  const figi = _getFigiByTicker(ticker)    // Tinkoff API v1 function !!!
  if (figi) {
    const data = tinkoffClientV2._GetLastPrices([figi])
    return Number(data.lastPrices[0].price.units) + data.lastPrices[0].price.nano/1000000000
  }
}

function TI_GetAccounts() {
  const data = tinkoffClientV2._GetAccounts()

  return data.accounts[0].id //FIXME!!!
}

function TI_GetPortfolio(accountId) {
  const portfolio = tinkoffClientV2._GetPortfolio(accountId)
  const values = []
  values.push(["FIGI","Название","Тип","Кол-во","Ср.цена покупки","Валюта","Доходность","Тек.цена","Валюта","НКД","Валюта"])
  for (let i=0; i<portfolio.positions.length; i++) {
    const [ticker,name] = _GetTickerNameByFIGI(portfolio.positions[i].figi)
    values.push([
      ticker,
      name,
      portfolio.positions[i].instrumentType,
      Number(portfolio.positions[i].quantity.units) + portfolio.positions[i].quantity.nano/1000000000,
      Number(portfolio.positions[i].averagePositionPrice.units) + portfolio.positions[i].averagePositionPrice.nano/1000000000,
      portfolio.positions[i].averagePositionPrice.currency,
      Number(portfolio.positions[i].expectedYield.units) + portfolio.positions[i].expectedYield.nano/1000000000,
      Number(portfolio.positions[i].currentPrice.units) + portfolio.positions[i].currentPrice.nano/1000000000,
      portfolio.positions[i].currentPrice.currency,
      Number(portfolio.positions[i].currentNkd.units) + portfolio.positions[i].currentNkd.nano/1000000000,
      portfolio.positions[i].currentNkd.currency
    ])
  }
  return values
}

