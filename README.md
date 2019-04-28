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

To set the risk free rate to the 1 month Daily Treasury Yield Curve Rate, use the **set_risk_free_rate_t_bill()** method:

```python
test_pf.set_risk_free_rate_t_bill()
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

## T-Bill rates

To load the most recent T-Bill rates, use the **get_T_bill_rates()** method:

```python
test_pf.get_T_bill_rates()
```

*The rates are stored in **self.t_bill_rates** as a pandas DataFrame object*

## Option prices 

To load available call and put option prices from Yahoo, use the **update_option_price(ticker)** and **update_option_price_list(tickers)** methods:

```python
ticker = "AAPL"
test_pf.update_option_price(ticker)

#Get option prices for the remaining stocks
tickers = test_pf.stocks[1:]
test_pf.update_option_price_list(tickers)
```

*An headless chrome driver is used via Selenium to access the Yahoo Option webpage. This is due to Beautiful Soup not being able to locate the drop-down menu options, which are needed to switch between avaible exercise dates. From there, the text values of the drop-down menu are converted into the proper Yahoo UNIX format, then used to build url request strings. Said strings are fed to Requests and Beautiful Soup before being converted to proper dataframes using Prandas' .read_html(...)*


# Extracting Call and Put Options data tables

Option data **__which has been previously downloaded__** can be extracted and visualised via the **get_call_put_data(ticker, expiration, option_type)** method:


```python
ticker = "AAPL"
expiration = "2019-05-24"
option_type = "Call"

AAPL_Call_DataFrame = test_pf.get_call_put_data(ticker, expiration, option_type)
print(AAPL_Call_DataFrame)
```

*Call and Put have to be capitalised.The option with the closest exercise date will be returned if no matches are found.*


# Purchasing Call and Put Options

To purchased call and put options based on avaible data, use the **buy_call_put_from_data(ticker, expiration, strike, n_purchase, option_type)** method:

```python
ticker = "AAPL"
expiration = "2019-05-24"
strike = 217.5
n_purchase = 12
option_type ="Call"

test_pf.buy_call_put_from_data(ticker, expiration, strike, n_purchase, option_type)

```

*Call and Put have to be capitalised.The option with the closest exercise date will be returned if no matches are found. A cubic spline is used to model Ask = f(Strike) within the data's range.*

# Exercising Call and Put Options

To exercise call and put options, use the **exercise_put_call()** method:


```python
test_pf.exercise_put_call()

```

*The current date is used to check if any options should have been exercised since last checked. Yhaoo daily High / Low prices are used to determine if options should have been exercised on their respective expiration dates. Said prices are also used to purchase missing shares if need be (i.e.: Put).*



```python

```







