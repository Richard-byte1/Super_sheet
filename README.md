# Supersheet

**Supersheet** is a Python-based toolkit that integrates with Excel, enabling advanced financial analytics and trading directly from your spreadsheets. Use it to run Monte Carlo simulations, calculate Value at Risk (VaR), pull and analyze financial data, place stock and option trades, and perform Discounted Cash Flow (DCF) analysis.

## Table of Contents


- [Funda](#Funda)
- [DCF](#DCF)
- [Portfolio and Stocks](#Portfolio_and_Stocks)
- [MC_Stock_Returns](#MC_Stock_Returns)
- [Connecting with Interactive Brokers](#Connecting_with_Interactive_Brokers)
- [BSM](#BSM)
- [Placing_Option](#Placing_Option)
- [Placing_Stock](#Placing_Stock)
- [Author and Contact](#Author_and_Contact)
- [Acknowledgments](#Ackowledgments)
- 
  ## Funda
  Using the Excel workbook and the "Final data" tab, you will be able to use the Fundamentals tab. This code uses "Alpha Vantage" to pull financial statements; you can get a free API key with an email address.
  
<img width="530" alt="image" src="https://github.com/user-attachments/assets/28b7e5d3-9c1f-44ef-8063-a721ea9459c2" />

Type in the stock to be looked up in E2 and, using the Excel "Stocks" function,

<img width="1405" alt="image" src="https://github.com/user-attachments/assets/5d790036-6275-4e9b-b9e5-51cb3059b237" />

the data should populate.
In the left-hand column, you should be able to see chart data, earnings estimates, ownership breakdown, and insider buying.

<img width="195" alt="image" src="https://github.com/user-attachments/assets/ff744ed5-d632-427e-ace3-34476e360286" />

The data is pulled into the Annual, Quarterly, and Ticker tabs.

  ## DCF

  In the DCF tab, you can create a DCF and, more importantly, sensitivity analysis. Information is used from the Funda tab. Switch between conservative, base, and optimistic cases.
WACC is the Weighted Average Cost of Capital, and TGR is the Terminal Growth Rate.

<img width="1282" alt="image" src="https://github.com/user-attachments/assets/47163539-1d92-47d4-b435-68044f475be9" />

On the WACC tab, you can calculate WACC.

<img width="617" alt="image" src="https://github.com/user-attachments/assets/2f215aaf-4b9e-49d9-ad25-55e3360079d8" />

You can use your own estimate or refer to https://pages.stern.nyu.edu/~adamodar/ for estimates of risk premiums.
Please use the interest rate on bonds for the cost of debt if available.
You can add other risk premiums if needed.

Please use the Revenue Step tab to better forecast your own estimates of future revenue.

## Portfolio_and_Stocks

To begin, add the stocks you bought under the stock column, along with the number of shares and the purchase price.

<img width="1467" alt="image" src="https://github.com/user-attachments/assets/9ce4d580-3a47-4a9a-af5d-6af358c1c591" />

This will update the Portfolio tab.
To run VaR, go to the VAR Inputs tab.

<img width="946" alt="image" src="https://github.com/user-attachments/assets/f5916824-1ca4-4ae0-b2ac-90f86ddbee8f" />

The end date is set to today, and the start date is set to -10 years.
The inputs are the number of simulations (normally 10,000), confidence interval, and the number of days to run the period over.
Run the code.
The code uses yfinance to pull data from Yahoo Finance.

## MC_Stock_Returns

Similarly, this uses Yahoo Finance to simulate stock returns over a specific time period.
The end date by default is today().
Please note there are 252 trading days in a year.
Threshold is the amount to be achieved.

<img width="232" alt="image" src="https://github.com/user-attachments/assets/e0252256-9326-47e3-bd29-164220fe8dfb" />

In the output, you can see the annualized mean return, annualized volatility, and the probabilities of reaching the threshold.

<img width="229" alt="image" src="https://github.com/user-attachments/assets/8976cc68-c1a3-435d-8ea4-a11a5ab93f5f" />

You can also see the frequency and the paths the simulations took.
Please note the outcome is heavily biased to the historical period chosen.

<img width="580" alt="image" src="https://github.com/user-attachments/assets/3cd4f6b4-272d-4211-915e-b26835724d82" />

<img width="846" alt="image" src="https://github.com/user-attachments/assets/c80e2790-c915-4ade-a64f-e21814e93df3" />

## Connecting with Interactive Brokers
Create an IBKR account and download Trader Workstation.
Follow IBKR’s official YouTube playlist on how to set up the IBKR API:
https://www.youtube.com/watch?v=i_5R8gkVAw0&list=PL71vNXrERKUpPreMb3z1WGx6fOTCzMaH1

Please use the following to see what is available using their API:  
https://interactivebrokers.github.io/tws-api/options.html

Options data requires a paid market subscription for personal use (15 USD per month).

TWS must be running in the backround

## BSM
In this tab, the inputs are the ticker, strike price, call or put, and risk-free rate. Running the code will allow you to see the Greeks and implied leverage for the option.

For a call, N(d2) is the risk-neutral probability that the option expires in-the-money.
For a put, N(-d2) is the risk-neutral probability that the option expires in-the-money.

Greeks can be used for options strategies discussed later.
Please note (I do not include Rho).

## Placing_Option

Place the ticker, sec type (usually STK for "stock"), exchange (SMART), and currency.
Run the code to make sure you are working on the correct stock.
Run Option Details to see available expiration dates and strikes.

<img width="437" alt="image" src="https://github.com/user-attachments/assets/ac256632-8f3f-4251-acdb-95264b38ebd3" />

Finally, once you have selected the options strategy, run the code to place the trade.

<img width="261" alt="image" src="https://github.com/user-attachments/assets/f32fe4c0-6037-405e-8656-051ca0fc0452" />

## Placing_Stock

In this tab, you can place an order to buy or sell stock along with a stop-loss order.

<img width="211" alt="image" src="https://github.com/user-attachments/assets/129da103-1101-4381-87f8-a4eea5c9e8b5" />

## Author and Contact
Richard A. For questions or support, open an issue or contact [Richard-byte1](https://github.com/Richard-byte1)

## Acknowledgments
- [@nickderobertis](https://www.youtube.com/@nickderobertis) for Excel and Python
- *Python for Finance* — Yves Hilpisch (Book)
