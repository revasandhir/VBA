# customer_accounts.bas 

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

- # recent_sales.bas
- # Recent Sales VBA Project

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

## Requirements
- Microsoft Excel with VBA enabled.
- Correctly formatted state worksheets in the workbook.

