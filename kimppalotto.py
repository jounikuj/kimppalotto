# -*- coding: utf-8 -*-

import argparse
import datetime
import matplotlib.pyplot as plt
import pandas
import seaborn
import sys
import yfinance

def main(args):
    
    # Import data.
    df = get_data()
    
    # ... and plot it.
    plot(df, args.dpi, args.output)
    print("Script complete!")
    
    sys.exit()

def get_lottery_data():
    """
    Function imports the Excel file located in the data/ folder and
    matches it with OMXH25 stock market data

    Returns
    -------
    df : pandas DataFrame
        DataFrame that contains the lottery and OMXH25 market index data.
    """
            
    # Import Excel file as pd.DataFrame.
    print("Retrieving lottery data...")
    df = pandas.read_excel("data/data.xlsx", parse_dates=True)
            
    # Return pd.DataFrame.
    return df

def get_omxh25_data(start_date):
    """
    Retrieve up-to-date OMXH25 stock market index data with yfinance API.
    Returns market data as pandas DataFrame object. Returned DataFrame has
    two columns: omxh25_date and omxh25_price representing the closing
    price of stock after an indicated date.

    Parameters
    ----------
    start_date : datetime
        Start date for API request
    ticker : str, optional
        Code for retrieved stock market object. The default is "^OMXH25".

    Returns
    -------
    df : pandas DataFrame
        DataFrame that contains the up-to-date OMXH25 stock data.
    """
            
    # Import corresponding OMXH25 stock market index data.
    print("Retrieving OMXH25 stock market data...")
    df = yfinance.download(tickers="^OMXH25", start=start_date)
            
    # Exclude non-necessary data.
    df = df[["Close"]].reset_index()
            
    # Rename columns.
    columns = {
        "Date":"omxh25_date",
        "Close":"omxh25_price"
    }
            
    df = df.rename(columns=columns)
            
    # Return pd.DataFrame.
    return df
    
def get_data():
    """
    Merges lottery data and OMXH25 stock market index data. Returns merged
    data as pandas DataFrame object.

    Returns
    -------
    lottery_df : pandas DataFrame
        DataFrame that contains lottery data and matched OMXH25 data.
    """
            
    # Import lottery data.
    lottery_df = get_lottery_data()
            
    # Import OMXH25 stock market data.
    omxh25_df = get_omxh25_data(lottery_df["date"].min())

    # In Finland, lottery is drawn every Saturday which is not a stock
    # market day. To match lottery data with stock market data, we will
    # match lottery results with the closing price of next available
    # stock market day.

    print("Processing data...")
    market_dates = omxh25_df["omxh25_date"].tolist()

    for index, row in lottery_df.iterrows():
                
        # Extract lottery day.
        lottery_day = row["date"]
                
        # If extracted lottery day is not a stock market day, go through
        # the calendar until match with stock market data is found.
        while lottery_day not in market_dates:
                    
            lottery_day += datetime.timedelta(1)
                
            if lottery_day > market_dates[-1]:
                break
            
        # Add next available stock market day to lottery DataFrame.
        # This data will be later used in merging the DataFrames.
        lottery_df.at[index, "omxh25_date"] = lottery_day
                
        # Ensure that datetime data is in right format.
        lottery_df["omxh25_date"] = pandas.to_datetime(lottery_df["omxh25_date"])
                
    # Merge lottery and OMXH25 data.
    # Drop unnecessary columns.
    lottery_df = lottery_df.merge(omxh25_df, how="left", on="omxh25_date")
    lottery_df = lottery_df.drop(columns=["omxh25_date"])
            
    def calculate_lottery_win_loss(df):
        
        # Calculate cumulative win/loss for each round.
        df["lottery_price_cum"] = df["price"].cumsum()
        df["lottery_win_cum"] = df["win"].cumsum()
                    
        # Calculate relative win/loss for each round.
        df["lottery_win_rel"] = ((df["lottery_win_cum"] - df["lottery_price_cum"]) / df["lottery_price_cum"]) * 100
                    
        # Round results to one decimal accuracy.
        df["lottery_win_rel"] = round(df["lottery_win_rel"], 1)
                    
        # Return pd.DataFrame.
        return df
        
    def calculate_omxh25_win_loss(df):
                
        # Calculate how many stock one could have bought with the price
        # of lottery round.
        df["omxh25_units"] = df["price"] / df["omxh25_price"]
            
        # Calculate the cumulative number of stock units.
        df["omxh25_units_cum"] = df["omxh25_units"].cumsum()
                
        # Calculate the value of stocks after each lottery round.
        df["omxh25_value"] = df["omxh25_units_cum"] * df["omxh25_price"]
                
        # Calculate the relative calue development of OMXH25 stocks.
        df["omxh25_win_rel"] = ((df["omxh25_value"] / df["lottery_price_cum"]) - 1) * 100
        df["omxh25_win_rel"] = round(df["omxh25_win_rel"], 1)
                
        # Return pd.DataFrame.
        return df
            
    # Perform calculations.
    lottery_df = calculate_lottery_win_loss(lottery_df)
    lottery_df = calculate_omxh25_win_loss(lottery_df)
            
    # Reset index and rename columns.
    columns = {
        "index":"round"
    }
            
    lottery_df = lottery_df.reset_index().rename(columns=columns)
            
    # Return pd.DataFrame.
    return lottery_df

def plot(df, dpi, output):
    """
    Generates simple line plot from given pandas DataFrame. Generated
    figure is saved as image file to repository root.

    Parameters
    ----------
    df : pandas DataFrame
        DataFrame that contains processed lottery and OMXH25 data.
    dpi : int
        DPI of output image, default: 300 DPI

    Returns
     -------
    None, saves image as PNG file to repository root.
    """
            
    # Initialize new matplotlib objects.
    print("Plotting data...")
    fig, ax = plt.subplots(1, 1, figsize=(6, 4), dpi=100)
            
    # Generate plots.
    seaborn.lineplot(x="round", y="lottery_win_rel", data=df, ax=ax)
    seaborn.lineplot(x="round", y="omxh25_win_rel", data=df, ax=ax)
    seaborn.despine(offset=10)
            
    # Define plot title.
    latest_round = df["date"].max().strftime("%d.%m.%Y")
    ax.set_title("Porukkaloton tilanne {}".format(latest_round), fontweight="bold")

    # Define axis labels.
    ax.set_xlabel("Lottokierros", fontweight="bold")
    ax.set_ylabel("Voittoprosentti (%)", fontweight="bold")

    # Define y-axis limits.
    # Add horizontal line to y-axis zero point.
    ax.set_ylim(-100, 100)
    ax.axhline(0, color="black", linestyle=":", linewidth=1)
            
    # Define legend.
    def get_legend_values(df):
        """
        Converts the latest data values to custom legend labels that are
        used in plotting.

        Parameters
        ----------
        df : pandas DataFrame
            DataFrame that contains processed lottery and OMXH25 data.

        Returns
        -------
        legend_text : list of str
            List that contains custom legened labels
        """

        # Extract latest data values (relative win/loss)
        lottery = df["lottery_win_rel"].iloc[-1]
        omxh25 = df["omxh25_win_rel"].dropna().iloc[-1]
            
        # If either one of them is larger than zero, add a plus
        # sign to it (e.g. +2.5%),
        if lottery > 0:
            lottery = "+" + str(lottery)
                        
        if omxh25 > 0:
            omxh25 = "+" + str(omxh25)

        # Generate legend labels.    
        legend_text = [
            "Porukkalotto {}%".format(lottery),
            "OMX Helsinki {}%".format(omxh25)
        ]
            
        # Return labels.
        return legend_text

    ax.legend(get_legend_values(df), frameon=False, loc="best")
            
    # Save figure.
    plt.tight_layout()
    plt.savefig(output, bbox_inch="tight")

if __name__ == "__main__":
    
    # Define ArgumentParser.
    parser = argparse.ArgumentParser()
    
    # Define command line arguments.    
    parser.add_argument("-dpi", dest="dpi", type=int, default=300,
                        help="Image DPI, default 300 DPI")
    parser.add_argument("-o", dest="output", type=str, default="kimppalotto.png",
                        help="Name of output image file")
    
    args = parser.parse_args()
    main(args)