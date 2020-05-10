# gas-tinkoff-trades
Данный [Google Apps Script](https://developers.google.com/apps-script) предназаначен для импорта сделок из Тинькофф Инвестиций прямо в Google таблицы, для последующего ручного анализа. 

Я сделал эту программу для автоматизизации ручного вбивания данных из приложения тинькофф, и надеюсь она окажется полезной кому-нибудь ещё.


## Установка

* Создайте новый или откройте документ Google Spreadsheet http://drive.google.com
* В меню "Tools" выберите "Script Editor"
* Дайте проекту имя, например TinkoffTrades
* Скопируйте код из [Code.GS](https://raw.githubusercontent.com/ErhoSen/gas-tinkoff-trades/master/Code.gs)
* Получите [OpenApi-токен тинькофф](https://tinkoffcreditsystems.github.io/invest-openapi/auth/)
* Добавьте свойтво OPENAPI_TOKEN в разделе File -> Project properties -> Script properties равную токену, полученному выше. 
* Сохраните скрипт

На этом всё. Теперь при работе с этим документом на всех листах будут доступны 2 новые функции `getPriceByTicker` и `getTrades`.

## Особенности

* Скрипт зарезервировал ячейку `Z1` (самая правая ячейка первой строки), в котороую вставляется случайное число на каждое изменении листа. Данная ячейка используется в функции `getPriceByTicker`, - она позволяет [автоматически обновлять](https://stackoverflow.com/a/27656313) текущую стоимость тикера при обновлении листа.

* Среди настроек скрипта есть `TRADING_START_AT` - дата начиная с которой фильтруются операции `getTrades`. По умолчанию это `Apr 01, 2020 10:00:00`, но данную константу можно в любой момент поменять в исходном коде.

## Функции

* `=getPriceByTicker(ticker, dummy)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опциональный параметр `dummy`. Для автоматичекого обновления необходимо указать в качестве `dummy` ячейку `Z1`. 

* `=getTrades(ticker, from, to)` - требует на вход [тикер](https://ru.wikipedia.org/wiki/%D0%A2%D0%B8%D0%BA%D0%B5%D1%80), и опционально фильтрацию по времени. Параметры `from` и `to` типа datetime и должны быть в [ISO 8601 формате](https://ru.wikipedia.org/wiki/ISO_8601)

## Пример использования 

```
=getPriceByTicker("V", Z1)  # Возвращает текущую цену акции Visa
=getPriceByTicker("FXMM", Z1)  # Возвращает текущую цену фонда казначейских облигаций США

=getTrades("V") 
# Вернёт все операции с акцией Visa, которые произошли начиная с TRADING_START_AT и по текущий момент.
=getTrades("V", "2020-05-01T00:00:00.000Z") 
# Вернёт все операции с акцией Visa, которые произошли начиная с первого мая и по текущий момент.
=getTrades("V", "2020-05-01T00:00:00.000Z", "2020-05-05T13:00:00.000Z") 
# Вернёт все операции с акцией Visa, которые произошли в период с 1 и по 5 мая.
```


## Пример работы

![getTrades in action](https://github.com/ErhoSen/gas-tinkoff-trades/raw/master/images/get-trades-in-action.gif "getTrades in Action")
