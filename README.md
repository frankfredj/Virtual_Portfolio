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
tickers = ["AAPL, "SasaAsa", "BO", "BA", "CN", "WM", "TLSE", "MCS", "TMS", "GOOG", "AMZN", "DIS", "NFLX", "AIR.PA", "LMT"]

test_pf.add_stock(tickers)

```







