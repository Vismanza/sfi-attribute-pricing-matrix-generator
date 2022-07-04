# sfi-attribute-pricing-matrix-generator
Google Apps Script to run with a Google Sheet that will generate a sample attribute based pricing matrix

## How to use

Once the script is authorised on your sheet you will have a new menu item under 'Salesforce Industries' called 'Generate ABP Matrix'. Clicking on the menu item will fire a function that will loop through every possible combination of attributes values for each product listed. It will then complete some basic random price and costs for the products based on the attribute value combinations.

It is very easy to exceed Sheets limits of more 1M rows so set your expectations accordingly. Once the data is processed it will be put on the ABP Data tab and some basic summary info will be logged on the 'Stats' tab.

At this point you can just download the 'ABP Data' tab as a csv from the 'File' menu in Sheets.

Example [Sheet]([url](https://docs.google.com/spreadsheets/d/1l-b3gRsfO4GduV6umusO6YHpUKdF45sxoooVaE4QgxY/edit?usp=sharing)).

## Google Sheet Format

You will need a sheet with tabs called:
- Products
- Attributes
- ABP Data
- Stats

### Products Tab

The products tab will have one row per product with the relevant pricing configurations. Note that currently any prices not needed must be entered as 0 and margins as 'None'.

Needs the following headers (first row values):
- Product Name
- Product Code
- Charge Measurement (ensure that the values exists in your envrionment)
- Usage Min (lower bound of usage price)
- Usage Max (upper bound of usage price)
- Usage Margin (None | Low | Med | High)
- Recurring Min (lower bound of recurring price)
- Recurring Max (upper bound of recurring price)
- Recurring Margin (None | Low | Med | High)
- One Time Min (lower bound of one time price)
- One Time Max (upper bound of one time price)
- One Time Margin (None | Low | Med | High)

### Attributes Tab

This tab will have a column per attribute you are wanting to generate prices for. The format is that the Attribute name is the first row and the attributes values are listed in the column below. This allows a different number of attributes to be stored for each attribute. 

The function will work regardless of the number of attributes (subject to Sheets row limits) as well as any spaces. You can delete certain values in a column. The function will automically skip over any empty cells.

### ABP Data Tab

This will be cleared each time the function is invoked so it need only exist.

### Stats

Needs the following headers (first row values):
- Timestamp
- Attributes Used
- Combinations Priced
- Execution Time (seconds)
- Attributes


