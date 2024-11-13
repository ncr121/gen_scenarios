# Credit Scenario Engine
This project is a designed for helping traders to manage credit index options portfolio in multi-strategy hedge funds. The scenario engine runs sensitivity analysis by calculating delta and gamma for individual trades and the overall portfolio for small and extreme shifts in spreads.
## Description
For each trade, the bumped index levels are computed based on the input parameters provided for the trade's credit index. The scenario engine calls the Bloomberg CDSO pricer to compute a table where the rows are the greeks and the columns are the bumped index levels. This table shows how the greeks will change based on different scenarios (bumps).

Originally, this was done by calling the Bloomber-Python API directly, but due to latency issues based on the large number of trades and scenarios a more efficient method was used. This is by using the _xlwings_ package to write the formulas to Excel using Python then allow Excel to calculate automatically and reading the data back to Python.

Based on this data, delta and gamma plots are created as well as payoff diagrams at maturity against index level. For each trade, a new sheet is created containing the greek table and graphs.

Finally, for each credit index, the greeks and payoffs are aggregated and all tables and graphs are displayed in the _Summary_ sheet.
## Getting Started
### Configuration
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
### Running the code
Once the paramters are set, the file _gen_scenarios.py_ can be run and all tables and graphs will be generated on new sheets in _CDSO_deal_list.xlsx_.
