## files 
- [customer_accounts.xlsm]([https://1drv.ms/x/c/9ea28cb2b29586b0/Ed90ESMu-U5Ktp8I-koOJA8BZmGDDlWjglwudDV0oTD6bw?e=d82T7T](https://1drv.ms/x/c/9ea28cb2b29586b0/EWtiBmq8q8FOnPE-rK9sSyMBMKP_1UcMRL1OL1R6atJ4oA?e=6s2Rdn))
- [recent_sales.xlsm]([https://1drv.ms/x/c/9ea28cb2b29586b0/EXTYaXxKlOFIsXnogE9nbkEBy4yEfl0iD833HfC1QDpqkw?e=PZb4Aj](https://1drv.ms/x/c/9ea28cb2b29586b0/EUjRCnziqZJHs0Jet7rvEgUBjF8QfQnZjbhR7F0yxQUHtg?e=c1oftW))

## 1. customer_accounts.bas
   
This script identifies customers with outstanding balances greater than $1,000 and outputs the results to the "Results" sheet.

## Description:

The script dynamically processes customer account data from the "Accounts" sheet, which includes:
- Customer ID
- Dollar value of purchases
- Dollar amount paid

## Features:
- Clears any previous output from the "Results" sheet.
- Lists customers with balances greater than $1,000.
- Adapts to changes in the "Accounts" data (e.g., addition or removal of rows).

## How to Use
1. Open the VBA editor in Excel.
2. Copy the code into a module.
3. Ensure the "Accounts" and "Results" sheets are formatted as expected.
4. Run the script to view results.

## 2. recent_sales.bas
   
This VBA project replicates the functionality of the "Recent Sales Finished" file, allowing users to summarize sales by sales reps in various states.

## Features
- Prompts for a sales rep and state.
- Searches state-specific worksheets for sales data.
- Displays total sales for the sales rep or an error if the input is invalid.

## Usage
1. Open `Recent Sales.xlsm` in Excel.
2. Enable macros.
3. Run the `Main` subroutine (via `Alt + F8` or a button).
4. Enter a sales rep and state when prompted.

## Example Outputs
- **Valid Input**: `Total sales for John Doe in California: $25,400`
- **Invalid Sales Rep**: `Sales rep Jane Smith not found.`
- **Invalid State**: `State Alaska not found.`
