 
Scenario engine for helping traders manage credit index options portfolio in multi-strategy hedge funds - ran sensitivity analysis - calculated delta and gamma at trade level and overall portfolio for small and extreme shifts in spreads

Modular code that can easily be extended to run scenarios on other greeks like theta and vega.
Code reads in config file (_"CDSO_deal_list.xlsx"_)- an Excel file that contains all input parameters for trades (Bloomberg deal IDs from CDSO screen), type of greeks and size/# of bumps  All scenarios and plots are generated based on this config file.

Originally called Bloomberg CDSO pricer directly via Python API but given latency issues due to the large # of trades and scenarios needed, improved efficiency by using Python to paste the BDP function call into Excel and then simultaneously reprice all scenarios in Excel using a simple refresh calc. 

Created delta and gamma plots as well as the payoff diagrams at maturity against index level, all in Excel to allow interactive use for trading desk. Formatting for plots is done in _"xlwings_functions.py"_ using VBA syntax.





