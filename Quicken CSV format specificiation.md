Import transactions from .CSV file
You can import financial transactions into Quicken from a comma-separated values (.CSV) file. This feature is useful when you need to add historical data, manage accounts manually, or bring in transactions from a financial institution that isn't supported for download.

Before you begin
The .CSV file must follow a strict format. The file must include all required columns, in the correct order, even if some columns have no data. A header row is optional, but if used, it must match the expected column structure.

Required file format
The file must contain the first seven comma separated columns in exactly this order to be processed. The last three comma-separated columns are optional, but if used, should appear in this order:

Date *

Payee *

FI Payee

Amount *

Debit/Credit

Category

Account *

Tag

Memo

Chknum

Columns marked with * must contain data in each row. The other columns may be blank, but each column position must be present in the file.

Additional Notes
Date values must follow the Quicken mm/dd/yyyy format. Leading zeros are optional for single-digit months and days. For example, 5/31/2025.

Amount values may optionally include commas. For example, 5400 or 5,400 are both valid amounts.

The FI Payee column is included for backward compatibility with CSV files exported from Mint. It is ignored during import but must be included as a column.

The Debit/Credit column is also for Mint compatibility. We recommend leaving this column blank and using a signed value in the Amount column instead. A positive amount (for example, +5,400) increases net worth; a negative amount (for example, –5400) decreases it. You can use a plus (+) or minus (–) sign to indicate the transaction type. If there is no sign, Quicken assumes the amount is positive.

If a field contains a comma, enclose it in double quotes—for example, "Office Supplies, Inc."

Do not add, remove, or reorder columns.

A header row is optional. If you include one, it must follow the same column order.

Import the file
To import a properly formatted .CSV file:

Open your Quicken data file.

Select File > File Import > Transactions from .CSV File.

In the Import transactions from .CSV file window, review the format requirements.

Select Choose .CSV File to Import….

Browse to your file and select it.

Select OK.

If the file does not meet the required structure, an error message will explain what needs to be fixed.

When account names don’t match
If the file contains transactions for an account that doesn't exist in your data file, Quicken will prompt you to:

Ignore the transactions for that account,

Add them to an existing account, or

Create a new account and import the transactions into it.

Choose the appropriate option and continue the import.

After the import
When the import completes, Quicken displays a summary showing:

The number of transactions imported

Any new accounts created

Any tags added

To update the balance of a newly created or existing manual account:

Open the account register.

Select the gear icon.

Choose Update Balance.

Additional notes
The import process does not detect or create transfers automatically.

Reimporting the same .CSV file will create duplicate transactions.

Use a plain-text editor or spreadsheet tool that supports UTF-8 encoding to prepare your file.


Copyright © 2025 Quicken • Powered by Scroll Viewport & Atlassian Confluence

