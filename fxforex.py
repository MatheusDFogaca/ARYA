"""
File related to the program functions
used in the main GUI of the program
"""

# Import section
import yfinance as yf
import matplotlib.pyplot as plt
import pandas as pd
import os
import numpy as np
import datetime as dt



def get_dataframe(ticker, start='2000-01-01', end=None):
    """
    Gets a dataframe based on parameters set
    and returns a simple dataframe for the analysis
    This is used in other functions

    """
    ticker = ticker.split(',')
    lista = []
    for i in ticker:
        lista.append(i + '=X')

    if len(lista) != 1:
        df = yf.Tickers(lista).history(start=start, end=end)
    else:
        df = yf.Ticker(lista[0]).history(start=start, end=end)

    df.pop('Volume')
    df.pop('Dividends')
    df.pop('Stock Splits')
    df.dropna(axis='index', how='any')
    return df


def get_legenda(df):
    """Returns the legend of the dataframe to be used in the plot"""
    legenda_ruim = df.columns.to_list()
    legenda = []
    for i in legenda_ruim:
        try:
            legenda.append(i.replace('=X', ' '))
        except:
            legenda=legenda_ruim
            pass
    return legenda


def csv(ticker, start='2000-01-01', end=None):
    """ returns a csv file to the desktop of the user"""
    df = get_dataframe(ticker, start, end)
    df.to_csv(os.environ['userprofile'] + '\\Desktop\\FOREX_EXPORT.csv')


def plot(df, style=None, lw=1, fs=18, ticker=None, xlabel="Quote", cmap=None, bar=bool, color=int,cor=None,compare=False):
    """
    Creates and displays a matplotlib chart in the screen with the data inside the dataframe

    :param df: a Pandas Dataframe with index as date and at least a columns called 'Close' which is the closing rate of the day
    :param style: The matplotlib desired style to plot in the chart
    :param lw: Line width parameter, 1 is the smallest
    :param fs: Font size starts at 18, can be changed to better accommodate the labels
    :param ticker: currency input in the tkinter entry
    :param xlabel: Xlabel if used as Quote or % change it depends the program will change
    :param cmap: Color map, used to define a collor palette to plot
    :param bar: True or False, telling if it should be a bar graph
    :return: plots a chart into the screen
    """
    plt.style.use('default')
    plt.style.use(style)
    plt.rcParams["font.family"] = "EMprint"

    fig, ax = plt.subplots(figsize=(11, 5), dpi=100)
    if color ==1:
        pass
    elif color == 2:
        cmap_factor = len(ticker.split(',')) + 1
        colors = plt.get_cmap(cmap)
        colors = colors(np.linspace(.8, .2, cmap_factor))
        ax.set_prop_cycle('color', colors)
    else:
        ax.set_prop_cycle('color', cor)

    if bar == False:
        if compare == False:
            ax.plot(df['Close'], lw=lw)
        else:
            ax.plot(df.index,df,lw=lw)
            ax.legend(handles=ax.lines, labels=get_legenda(df), loc='best')
        ax.set_ylabel(xlabel, fontsize=fs)
        ax.set_xlabel("Period", fontsize=fs)

        if len(ticker.split(',')) != 1:
            legenda = get_legenda(df['Close'])
            ax.legend(handles=ax.lines, labels=legenda, loc='best')
        else:
            pass

        for i in (ax.get_xticklabels() + ax.get_yticklabels()):
            i.set_fontsize(fs / 1.5)

        if len(ticker.split(',')) > 1:
            ax.set_title(label=str(ticker) + ' vs. USD', fontsize=fs * 1.5, pad=15)

        elif len(ticker) > 3:
            ax.set_title(label=ticker[:3] + ' vs. ' + ticker[3:], fontsize=fs * 1.5, pad=15)
        else:
            ax.set_title(label=ticker + ' vs. USD', fontsize=fs * 1.5, pad=15)

    else:
        ax.bar(df.index , df['Close'], lw, 0)
        ax.set_ylabel(xlabel, fontsize=fs)
        ax.set_xlabel("Period", fontsize=fs)
        ax.set_title(label=ticker + ' vs. USD Daily Changes %', fontsize=fs * 1.5, pad=15)

    plt.show()


def get_base_100(df):
    for i in df['Close']:
        df['% ' + i] = df['Close'][i] / df['Close'][i].iloc[0] - 1

    df = df[-2:]
    return df


def get_LDM(df):
    df.reset_index(inplace=True)
    new_df = df.loc[df.groupby(pd.Grouper(key='Date', freq='1M')).Date.idxmax()]
    new_df = new_df.set_index('Date')
    return new_df


def pct_change(df):
    new_df = df.pct_change()
    return new_df


def get_year_comparison(df):
    df['Year'] = df.index.year
    df['Month'] = df.index.month
    df['Day'] = df.index.day

    pt = df.pivot_table(columns='Year', index=['Month', 'Day'])['Close']
    pt = pt[pt[pt.columns[-1]].notnull()].dropna(axis='index', how='any')
    pt = pt / pt.iloc[0]
    pt.reset_index(inplace=True)
    pt.pop('Month')
    pt.pop('Day')

    return pt


def plot_png(df, lw=1, fs=18, ticker=None, xlabel="Quote"):
    """
    Creates and displays a matplotlib chart in the screen with the data inside the dataframe

    :param df: a Pandas Dataframe with index as date and at least a columns called 'Close' which is the closing rate of the day
    :param style: The matplotlib desired style to plot in the chart
    :param lw: Line width parameter, 1 is the smallest
    :param fs: Font size starts at 18, can be changed to better accommodate the labels
    :param ticker: currency input in the tkinter entry
    :return: plots a chart into the screen
    """

    plt.style.use('classic')
    plt.rcParams["font.family"] = "EMprint"

    fig, ax = plt.subplots(figsize=(11, 5), dpi=80)
    last_month = df[df.index>df.index[-22]]
    ax.plot(df['Close'], lw=lw, color='#8EB7D7')
    ax.plot(last_month['Close'], lw=lw+1, color='#0A4776')
    #ax.legend(handles=ax.lines, labels=get_legenda(df), loc='best')
    ax.set_ylabel(xlabel, fontsize=fs)
    ax.set_xlabel("Period", fontsize=fs/1.5)
    plt.xticks(rotation=25)

    for i in (ax.get_xticklabels() + ax.get_yticklabels()):
        i.set_fontsize(fs / 1.5)

    if len(ticker.split(',')) > 1:
        ax.set_title(label=str(ticker) + ' vs. USD', fontsize=fs * 1.5, pad=15)

    elif len(ticker) > 3:
        ax.set_title(label=ticker[:3] + ' vs. ' + ticker[3:], fontsize=fs * 1.5, pad=15)
    else:
        ax.set_title(label=ticker + ' vs. USD', fontsize=fs * 1.5, pad=15)



    plt.savefig(f"{os.environ['USERPROFILE']}\\plot.png",dpi=80)
    last_month.to_csv(f"{os.environ['USERPROFILE']}\\quotes.csv")