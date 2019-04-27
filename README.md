# Constructor inputs

**directory** (string) : OS path to directory

**file_name** (string) : file name within the directory

```python
test_pf = Portfolio(r"C:\Users\Francis\SkyDrive\Documents", "test_pf_2")
```

*If the provided path belongs to an already existing file, the contructor will load the existing portfolio. Otherwise, it will create
a new xlsx document to save portfolio informations.*


# Unlocking stocks

By default, a Portfolio object doesn't contain any stocks. Its .xslx file wasn't initated with multiple pages for each available commodity
on the market (obviously). To create new sheets that will be able to store data pertaining to specific stocks, use the **add_stock(tickers)** method:

```python
tickers = ["AAPL", "SasaAsa", "BO", "BA", "CN", "WM", "TLSE", "MCS", "TMS", "GOOG", "AMZN", "DIS", "NFLX", "AIR.PA", "LMT"]
test_pf.add_stock(tickers)
```

# Purchasing stock shares

There are two ways to purchase shares once you've unlocked stocks. You can either pass a single ticker name along with the desired number of shares with the **add_single_shares(ticker, amount)** method, or pass a vector to the **add_vector_shares(amounts)** method. 

*If using the vectorised method, one must provide a vector of length equal to the number of unlocked stocks.*


```python
test_pf.add_single_shares("AAPL", 5)
test_pf.add_vector_shares([15,12,19,5,6,13,15,12,19,5,6])
```

# Rates

The default borrowing rate is 0.05, whereas the default risk-free rate is 0.025. To modify these parameters, use the following **set_...(r)** methods:

```python
test_pf.set_borrowing_rate(0.039)
test_pf.set_risk_free_rate(0.241)
```

# Updating the Portfolio

To update the value of your stock equity and update your debt with respect to the borrowing rate (effective annual converted to continuous - unit time in seconds), use the **update()** method:

```python
test_pf.update()
```

# Loading historical data

## Yahoo historical prices

To update the historical data spreadsheets accoding to a certain date range, use **update_historical_data(start_date, end_date)** method:

```python
test_pf.update_historical_data("2018-04-26", "2019-02-26")
```

*Date formating needs to be Y%-%m-%d*

To extend the historical data up to the current date, use the **bridge_historical_data()** method:

```python
test_pf.bridge_historical_data()
```

## Option prices (experimental)

Scrapping the Barchart.com website has proven to be too much of an hard task given my limited html knowledge. However, I've devised an efficient (but slow) way to obtain the price tables by using Selenium alongside the OS clipboard. It basically copy-pastes the webpage's text and retrieves the table by formating the resulting string. 

Two methods are currently avaible: one for single stocks (**store_option_prices(ticker, expiration)**) and one for stock lists (**store_option_prices_list(tickers, expirations)**).

```python
tickers = test_pf.stocks
expirations = []

for i in range(len(tickers)):
	expirations.append("2019-05-26")
  
test_pf.store_option_prices_list(tickers, expirations)
```

*Since only a handful of expiration dates are avaible, the closest date will be selected.*







