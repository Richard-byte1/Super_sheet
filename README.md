# Supersheet


## Table of Contents

- [Funda](#funda)
- [DCF](#dcf)
- [Portfolio and Stocks](#portfolio-and-stocks)
- [MC_Stock_Returns](#mc_stock_returns)
- [Connecting with Interactive Brokers](#connecting-with-interactive-brokers)
- [Options (Order Book, Payoffs, and Strategies)](#options-order-book-payoffs-and-strategies)
- [Placing Option](#placing-option)
- [Placing Stock](#placing-stock)
- [MarketSnapshot](#marketsnapshot)
- [Author and Contact](#author-and-contact)
- [Acknowledgments](#acknowledgments)

## Funda

Using the Excel workbook and the "Final data" script, you can access the Fundamentals tab. This code uses "Alpha Vantage" to pull financial statements; you can obtain a free API key by registering with your email address.

<img width="530" alt="image" src="https://github.com/user-attachments/assets/28b7e5d3-9c1f-44ef-8063-a721ea9459c2" />

Type the stock symbol you want to look up in E2 and, using the Excel "Stocks" function,

<img width="1405" alt="image" src="https://github.com/user-attachments/assets/5d790036-6275-4e9b-b9e5-51cb3059b237" />

the data should populate.
In the left-hand column, you can view chart data, earnings estimates, ownership breakdown, and insider buying.

<img width="195" alt="image" src="https://github.com/user-attachments/assets/ff744ed5-d632-427e-ace3-34476e360286" />

The data is pulled into the Annual, Quarterly, and Ticker tabs.

**Chart Update with RSI**  
This script is now in Final data:

<img width="1453" height="613" alt="image" src="https://github.com/user-attachments/assets/7db394fe-e389-4a94-bc88-6e2508578418" />

## DCF

In the DCF tab, you can create a DCF and, more importantly, perform sensitivity analysis. Information is sourced from the Funda tab. Switch between conservative, base, and optimistic scenarios.
WACC is the Weighted Average Cost of Capital, and TGR is the Terminal Growth Rate.

<img width="1282" alt="image" src="https://github.com/user-attachments/assets/47163539-1d92-47d4-b435-68044f475be9" />

On the WACC tab, you can calculate WACC.

<img width="617" alt="image" src="https://github.com/user-attachments/assets/2f215aaf-4b9e-49d9-ad25-55e3360079d8" />

You can use your own estimate or refer to [NYU Stern's estimates of risk premiums](https://pages.stern.nyu.edu/~adamodar/).
Please use the interest rate on bonds for the cost of debt if available.
You can add other risk premiums if needed.

Please use the Revenue Step tab to better forecast your own revenue estimates.

## Portfolio and Stocks

To begin, add the stocks you purchased under the stock column, together with the number of shares and the purchase price.

<img width="1467" alt="image" src="https://github.com/user-attachments/assets/9ce4d580-3a47-4a9a-af5d-6af358c1c591" />

This will update the Portfolio tab.
To run VaR, go to the VAR Inputs tab.

<img width="946" alt="image" src="https://github.com/user-attachments/assets/f5916824-1ca4-4ae0-b2ac-90f86ddbee8f" />

The end date is set to today, and the start date is set to -10 years.
The inputs are the number of simulations (normally 10,000), confidence interval, and the number of days to run the period over.
Run the code.
The code uses yfinance to pull data from Yahoo Finance.

### Portfolio Update

The portfolio now shows Conditional VaR in the portfolio, included in the VAR script.

<img width="603" height="562" alt="image" src="https://github.com/user-attachments/assets/6137267b-9ff3-4e1d-8b09-3f8e66519e39" />

### Portfolio Optimization and Backtest

In this section, you select the time period, max volatility, minimum rate of return, the tickers in the portfolio, and the weights.

<img width="973" height="643" alt="image" src="https://github.com/user-attachments/assets/4d04370d-9358-406f-bcf2-980eb4c2f52c" />

This will give you five optimal portfolios depending on your goals, as well as their weights.

On the right, you will see portfolio performance

<img width="998" height="505" alt="image" src="https://github.com/user-attachments/assets/055f06c5-c68b-4675-a669-b739da670c72" />

and portfolio metrics

<img width="591" height="645" alt="image" src="https://github.com/user-attachments/assets/5491a548-02fe-4d14-a94b-68fcf96495e4" />

<img width="537" height="637" alt="image" src="https://github.com/user-attachments/assets/af57d73a-c2ad-4cdf-aa8c-1306d825432b" />

## Regression (Five_Factors)

In this sheet, you can perform a five-factor regression on a stock. Note: ^GSPC is the SPX.

<img width="945" height="514" alt="image" src="https://github.com/user-attachments/assets/227e0100-3070-46e0-baf5-c640c8af2dc7" />

<img width="717" height="504" alt="image" src="https://github.com/user-attachments/assets/d8a68309-83c3-4887-9f5f-cab3c7b4019c" />

## MC_Stock_Returns

Similarly, this uses Yahoo Finance to simulate stock returns over a specific time period.
The end date by default is today().
Please note there are 252 trading days in a year.
Threshold is the amount to be reached.

<img width="232" alt="image" src="https://github.com/user-attachments/assets/e0252256-9326-47e3-bd29-164220fe8dfb" />

In the output, you can see the annualized mean return, annualized volatility, and the probabilities of reaching the threshold.

<img width="229" alt="image" src="https://github.com/user-attachments/assets/8976cc68-c1a3-435d-8ea4-a11a5ab93f5f" />

You can also see the frequency and the paths the simulations took.
Please note the outcome is heavily biased by the historical period chosen.

<img width="580" alt="image" src="https://github.com/user-attachments/assets/3cd4f6b4-272d-4211-915e-b26835724d82" />

<img width="846" alt="image" src="https://github.com/user-attachments/assets/c80e2790-c915-4ade-a64f-e21814e93df3" />

## Connecting with Interactive Brokers

Create an IBKR account and download Trader Workstation (TWS).
Follow IBKR’s official YouTube playlist on how to set up the IBKR API:  
https://www.youtube.com/watch?v=i_5R8gkVAw0&list=PL71vNXrERKUpPreMb3z1WGx6fOTCzMaH1

Please consult the [IBKR API documentation](https://interactivebrokers.github.io/tws-api/options.html) to see what is available.

Options data requires a paid market subscription for personal use (15 USD per month).

TWS must be running in the background.

## Options (Order Book, Payoffs, and Strategies)

The options section has been completely reworked for better usability.
In the top left, enter the ticker, strike, expiry date, and option type, then run the code. You will get info and the order book from Yahoo Finance.

<img width="297" height="231" alt="image" src="https://github.com/user-attachments/assets/a94d9a3f-37eb-4f9c-8637-62920cce8bee" />

<img width="1463" height="703" alt="image" src="https://github.com/user-attachments/assets/ea3cf59a-4fbd-4b2c-94fe-538358f0be7c" />

<img width="1151" height="706" alt="image" src="https://github.com/user-attachments/assets/2c252a55-8729-4946-83d7-a006e18f32a3" />


In the Opt_strat sheet, fill in the inputs in the top left to get the Greeks.

<img width="568" height="132" alt="image" src="https://github.com/user-attachments/assets/c8384fd9-1d05-4a2d-aa12-27a82d1c0b2e" />

Select the strategy:

<img width="152" height="260" alt="image" src="https://github.com/user-attachments/assets/d7435acc-c88e-4adf-83e0-f6d6b3de222d" />

The K inputs manage the legs of the option strategy:

<img width="149" height="98" alt="image" src="https://github.com/user-attachments/assets/d241b779-cc85-49f2-bee1-7f7def398af8" />

View the payoff chart, implied volatility, and implied volatility surface:

<img width="1252" height="445" alt="image" src="https://github.com/user-attachments/assets/670142d0-1727-469b-86bf-e8145976f844" />

<img width="1303" height="476" alt="image" src="https://github.com/user-attachments/assets/f6b04353-fa7d-41f5-bff0-809098a5d75d" />

Finally, in the Option_tws sheet, select the strategy:

<img width="463" height="296" alt="image" src="https://github.com/user-attachments/assets/e8e93ed1-cf74-4145-94ce-e2b64b606a7d" />

Fill in the inputs and run the code:

<img width="467" height="53" alt="image" src="https://github.com/user-attachments/assets/28ca5821-cf52-4e56-afd7-7326fa453c06" />

<img width="503" height="140" alt="image" src="https://github.com/user-attachments/assets/d9371218-9e74-43c8-ace7-31735fb2d681" />

<img width="748" height="123" alt="image" src="https://github.com/user-attachments/assets/1aa16f94-44fc-4e1a-9495-de0849fdba48" />

Please note, you must have the correct trading permission to use some strategies.

## Placing Stock

In this tab, you can place an order to buy or sell stock along with a stop-loss order.

<img width="211" alt="image" src="https://github.com/user-attachments/assets/129da103-1101-4381-87f8-a4eea5c9e8b5" />

## TWAP
The TWAP code doesn't connect with Excel, but must run in Terminal, inputs are at the top of the script, and it's the perfect way to average into and out of positions

## MarketSnapshot

Please note that you will need a FRED API, and a free TIKR account, you cant set it up to run locally at 9:00am (or any time). Sources are included.

<img width="1394" height="555" alt="image" src="https://github.com/user-attachments/assets/0e3369ad-35ad-46f2-aa95-e223dcae6a1b" />

<img width="763" height="437" alt="image" src="https://github.com/user-attachments/assets/28abea9e-38f1-4914-b181-4d9c7dd0b656" />

<img width="966" height="361" alt="image" src="https://github.com/user-attachments/assets/794be1fc-cf63-41ed-9368-b57e3b2e1dc8" />

<img width="504" height="467" alt="image" src="https://github.com/user-attachments/assets/e49a8b05-4d2c-4c14-83be-9e1163297895" />





## Author and Contact

Richard A. For questions or support, open an issue or contact [Richard-byte1](https://github.com/Richard-byte1).

## Acknowledgments

- [@nickderobertis](https://www.youtube.com/@nickderobertis) for Excel and Python inspiration
- *Python for Finance* — Yves Hilpisch (Book)
