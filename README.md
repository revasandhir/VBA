# q1_customer_accounts.bas

This script identifies customers with outstanding balances greater than $1,000 and outputs the results to the "Results" sheet.

## Description
The script dynamically processes customer account data from the "Accounts" sheet, which includes:
- Customer ID
- Dollar value of purchases
- Dollar amount paid

### Features
- Clears any previous output from the "Results" sheet.
- Lists customers with balances greater than $1,000.
- Adapts to changes in the "Accounts" data (e.g., addition or removal of rows).

## How to Use
1. Open the VBA editor in Excel.
2. Copy the code into a module.
3. Ensure the "Accounts" and "Results" sheets are formatted as expected.
4. Run the script to view results.

### Dependencies
- Microsoft Excel with VBA enabled
