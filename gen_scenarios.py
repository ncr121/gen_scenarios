import os
import time
import datetime as dt
import numpy as np
import pandas as pd
import xlwings as xw
from collections import defaultdict

import xlwings_functions as xwfn

pd.options.mode.chained_assignment = None


def get_index(issuer):
    """
    Get the index from the issuer field of the deal.

    :param issuer: Issuer of deal.
    :return index: Index of deal.
    """
    # looping through column H of 'Deal_List' sheet
    for index in bump_params:
        if index in issuer:
            return index


def get_bumps(up_bump, down_bump, num_up, num_down):
    """
    Generate a bump vector based on parameters.

    :param up_bump: Size of up bump.
    :param down_bump: Size of down bump.
    :param num_up: Number of up bumps.
    :param num_down: Number of down bumps.
    :return bumps: Bump vector.
    """
    up_bumps = np.arange(num_up + 1) * up_bump
    down_bumps = np.arange(-num_down, 0) * down_bump
    return np.concatenate([down_bumps, up_bumps])


def format_bump_table(cell, number_formats):
    """
    Format the cells of a bump table in Excel.

    :param cell: Top left cell of bump table.
    :param number_formats: Excel number formats to apply to each row of the table.
    :return:
    """
    # write header names and add formatting
    cell.value = [['Bumped Index Level'], ['Bump Amount']]
    cells = cell.expand('table')
    cells[:2, 0].font.italic = True
    # add formatting to index levels
    cells[0, 1:].number_format = '0.00'
    cells[0, 1:].font.italic = True
    # add formatting to bumps
    cells[1, 1:].font.bold = True

    data = cells[2:, 1:]
    # add formatting to data of bump table
    data.color = (255, 255, 0)
    data.api.Borders.Weight = 2
    for row, number_format in zip(data.rows, number_formats):
        row.number_format = number_format


def xlwings_plot(y_values, x_values, left, top, width=355, height=211, chart_type='line', title=None,
                 legend=False, x_label=None, y_label=None, x_number_format='General',
                 y_number_format='General'):
    """
    Create a chart in Excel. One row is about 14.5 points and once columns is about 47.5 points.

    :param y_values: Y-values of
    :param x_values:
    :param left: Left starting point of chart in points.
    :param top: Top starting point of chart in points.
    :param width: Width of chart in points.
    :param height: Height of chart in points.
    :param chart_type: Chart type.
    :param title: Title of chart.
    :param legend:
    :param x_label: Label of x-axis.
    :param y_label: Label of y-axis.
    :param x_number_format: Number format of x-axis.
    :param y_number_format: Number format of y-axis.
    :return:
    """
    chart = ws.charts.add(left, top, width, height)
    chart.chart_type = chart_type
    # chart.set_source_data(cells)
    wrapper = chart.api[1]
    wrapper.HasLegend = legend
    series = wrapper.SeriesCollection().NewSeries()
    series.Values = y_values
    series.XValues = x_values

    if title is not None:
        wrapper.SetElement(2)
        wrapper.ChartTitle.Text = title

    x_axis, y_axis = (wrapper.Axes(i) for i in (1, 2))
    x_axis.TickLabels.NumberFormat = x_number_format
    y_axis.TickLabels.NumberFormat = y_number_format
    x_axis.TickLabelPosition = -4134

    if x_label is not None:
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = x_label

    if y_label is not None:
        y_axis.HasTitle = True
        y_axis.AxisTitle.Text = y_label

deal_wb = xw.Book('CDSO_deal_list.xlsx')
deal_ws = deal_wb.sheets['Deal_List']
# read in deal IDs, from column A, as a list
deal_ids = deal_ws['A2'].expand('down').value
# read in static fields, from in column C, as a list
static_fields = deal_ws['C2'].expand('down').value
# read in greeks and their respective number formats, from columns E:F, as a dict
greeks = deal_ws['E2'].expand('table').options(dict).value
# read in index types and their respective bump parameters, from columns H:L, as a dict
bump_params = deal_ws['H1'].expand('table').options(pd.DataFrame).value.T.to_dict()

agg_dfs = {}
agg_payoffs = {}
greek_tables = defaultdict(dict)

# fname = 'CDSO_live_trades_{}.xlsx'.format(dt.date.today())
# wb = xwfn.new_book(fname)
wb = deal_wb
for ws in wb.sheets:
    if ws.name != 'Deal_List':
        ws.delete()

for deal_id in deal_ids:
    ws = xwfn.new_sheet(deal_id, wb)
    # set centre alignment
    ws.cells.api.HorizontalAlignment = -4108
    # write deal ID to cell A1 and add colour and bold formatting
    ws['A1'].value = deal_id + ' Corp'
    ws['A1'].color = (146, 208, 80)
    ws['A1'].font.bold = True

    # write static field names to column A and their formulas to column B
    ws['A3'].value = [[field] for field in static_fields]
    ws['A3'].expand('down').offset(0, 1).formula = '=@BDP($A$1,A3)'
    time.sleep(5)
    # read in static fields for the deal, from columns A:B, as a dict
    deal_fields = ws['A3'].expand('table').options(dict).value
    # format notional of deal with commas separating every thousand
    ws[static_fields.index('SW_PAY_NOTL_AMT') + 2, 1].number_format = '#,##'
    index = get_index(deal_fields['ISSUER'])

    bumps = get_bumps(**bump_params[index])
    index_levels = deal_fields['PX_LAST'] + bumps
    # write empty bump table to sheet
    ws['D1'].value = pd.DataFrame(0, greeks, pd.MultiIndex.from_tuples(zip(index_levels, bumps)))
    # write formulas to bump table
    ws['E3'].expand('table').formula = '=@BDP($A$1,$D3,"CDS_FLAT_SPREAD",E$1)'
    time.sleep(5)
    format_bump_table(ws['D1'], greeks.values())
    # read in bump table with values
    deal_df = ws['D1'].expand('table').options(pd.DataFrame, header=2).value
    ws.autofit()

    # calculate notional in millions
    notional_mil = deal_fields['SW_PAY_NOTL_AMT'] / 1e6
    # slice greeks and multiply Delta and Gamma by notional
    greek_table = deal_df.iloc[2:]
    greek_table.iloc[:2] *= notional_mil
    greek_tables[index][deal_id] = greek_table

    # read in whether the deal is long/short and payer/receiver
    long_short = deal_fields['SW_CS_POSITION']
    pay_rec = deal_fields['SW_CDS_BUY_SELL_FLAG']

    k_minus_s = deal_fields['SW_SPREAD'] - index_levels
    # k - s for receiver and s - k for payer (if it is a spread index like XO, Main, IG)
    intrinsic_value = (k_minus_s if pay_rec == 'REC' else -1 * k_minus_s) * (-1 if index == 'HY' else 1)
    dvo1 = 4
    abs_payoff = notional_mil * np.maximum(0, intrinsic_value) / 1e4 * (1 if index == 'HY' else dvo1)
    # multiply payoff by -1 if short
    payoff = abs_payoff if long_short == 'LONG' else -1 * abs_payoff

    # write payoff for each index level to sheet
    ws['D10'].value = 'PnL'
    ws['D10'].font.bold = True
    ws['E10'].value = payoff

    # aggregate
    if index in agg_dfs.keys():
        agg_dfs[index] += greek_table
        agg_payoffs[index] += payoff
    else:
        agg_dfs[index] = greek_table
        agg_payoffs[index] = payoff

    # plot Delta (remove hard coding)
    xwfn.xlwings_plot(ws, greek_table.loc['SW_OPTION_DELTA'], index_levels, 'A12', 'D25',
                  title='Delta in millions against Index Level', x_label='Index Level',
                  y_label='Delta in millions', x_number_format='0.0')

    # plot payoff (remove hard coding)
    xwfn.xlwings_plot(ws, payoff, index_levels, 'F12', 'O25', title='Payoff at Maturity against Index Level',
                  x_label='Index Level', y_label='Payoff in millions', x_number_format='0.0')

ws = xwfn.new_sheet('Summary', wb)
# set centre alignment
ws.cells.api.HorizontalAlignment = -4108
for i, (index, df) in enumerate(agg_dfs.items()):
    # write index to cell in column A above bump table and add colour and bold formatting
    ws[16*i, 0].value = index
    ws[16*i, 0].font.bold = True
    ws[16*i, 0].color = (146, 208, 80)
    # write aggregated bump table to index
    ws[16*i + 1, 0].value = df
    format_bump_table(ws[16*i + 1, 0], ['0.0', '0.0', '0', '0'])

for i, (index, df) in enumerate(agg_dfs.items()):
    # plot Delta for index
    xwfn.xlwings_plot(ws, df.loc['SW_OPTION_DELTA'], df.columns.get_level_values(0), (16*i, 15), (16*i + 14, 22),
                  title='Delta against Index Level for {} Index'.format(index), x_label='Index Level',
                  y_label='Delta in millions', x_number_format='0.0')

    # plot payoff for index
    xwfn.xlwings_plot(ws, agg_payoffs[index], df.columns.get_level_values(0), (16*i, 24), (16*i + 14, 31),
                  title='Payoff against Index Level for {} Index'.format(index), x_label='Index Level',
                  y_label='Payoff in millions', x_number_format='0.0')

ws = xwfn.new_sheet('Delta Summary', wb)
delta_col_list = ['Existing CDS', 'Current Delta Equiv', 'Notional Needed for Delta Neutral', 'Index Level']
delta_df = pd.DataFrame(0, index=agg_dfs.keys(), columns=delta_col_list)
for index, agg_df in agg_dfs.items():
    current_delta = agg_df.loc['SW_OPTION_DELTA'].xs(0, level=1)
    delta_df.loc[index, 'Current Delta Equiv'] = current_delta.iloc[0]
    delta_df.loc[index, 'Index Level'] = current_delta.index[0]
    delta_df.loc[index, 'Notional Needed for Delta Neutral'] = delta_df.loc[index, 'Current Delta Equiv'] - \
                                                               delta_df.loc[index, 'Existing CDS']

    ws['A1'].value = delta_df
    cells = ws['B2'].expand('table')
    cells.color = (255, 255, 0)
    cells.api.Borders.Weight = 2
    cells[:, :3].number_format = '0.0'
    cells[:, 3].number_format = '0.00'
    cells[:, 2].font.bold = True
    cells.api.HorizontalAlignment = -4108

ws.autofit()

wb.save()
