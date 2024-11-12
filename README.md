# Credit Scenario Engine
This project is a scenario engine for helping traders to manage credit index options portfolio in multi-strategy hedge funds.
## Description
The scenario engine runs sensitivity analysis by calculating delta and gamma for individual trades and the overall portfolio.  for small and extreme shifts in spreads

Modular code that can easily be extended to run scenarios on other greeks like theta and vega.
Code reads in config file (_"CDSO_deal_list.xlsx"_)- an Excel file that contains all input parameters for trades (Bloomberg deal IDs from CDSO screen), type of greeks and size/# of bumps  All scenarios and plots are generated based on this config file.

Originally called Bloomberg CDSO pricer directly via Python API but given latency issues due to the large # of trades and scenarios needed, improved efficiency by using Python to paste the BDP function call into Excel and then simultaneously reprice all scenarios in Excel using a simple refresh calc. 

Created delta and gamma plots as well as the payoff diagrams at maturity against index level, all in Excel to allow interactive use for trading desk. Formatting for plots is done in _"xlwings_functions.py"_ using VBA syntax.
## Getting Started
Go to _CDSO_deal_list.xlsx_ where all input parameters can be set in the config sheet titled _Deal_List_. Below are the descriptions for each type of input:

Column A (__Deal IDs__): Bloomberg deal IDs from CDSO screen for valid trades

Column C (__Static Fields__): Static fields to be read in from Bloomberg to for each trade

Column E (__Greeks__): Greeks to perform sensitivity analysis on at a trade and portfolio level

Column F (number_format): Excel number formatting for each greek respectively

Column G (__Index__): Credit indices

Column H (up_bump): Size of bumps in the positive direction

Column I (down_bump): Size of bumps in the negative direction

Column J (num_up): Number of bumps in the positive direction

Column K (num_down): Number of bumps in the negative direction
