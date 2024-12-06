# Customer Accounts Analysis

This VBA project identifies customers with outstanding balances over $1,000 and generates a report on the "Results" sheet. 

## Description

The "Accounts" sheet contains customer account data with the following fields:
- Customer ID
- Dollar value of customer purchases in the current year
- Dollar amount paid for these purchases

### Example
- **Customer 1302:** Purchased $2,466 worth of goods and has paid the full amount.
- **Customer 2245:** Purchased $1,494 worth of goods but has only paid $598, leaving an outstanding balance.

### Key Features
- The script outputs a list of customers who owe strictly more than $1,000.
- The results include both the customer ID and the dollar amount outstanding.
- Previous outputs on the "Results" sheet are cleared before generating new results.
- The code dynamically adjusts to changes in the "Accounts" data, such as adding or removing customers.

## How to Use
1. Place the code in the VBA editor of your Excel workbook.
2. Ensure the "Accounts" and "Results" sheets are properly named and formatted.
3. Run the script to generate the results.
