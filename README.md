# VenmoToQuicken

This PowerShell script converts Venmo transaction CSV files into a format compatible with Quicken Classic. Venmo does not natively support exporting transactions in a Quicken-compatible format, so this script bridges the gap by transforming the Venmo CSV format into the required Quicken CSV structure.

## Features

- Converts Venmo CSV files to Quicken-compatible CSV files.
- Automatically detects and skips balance summary lines.
- Supports custom account names and date formats.
- Handles missing or optional fields gracefully.
- Outputs a properly formatted CSV file ready for import into Quicken.

## Requirements

- PowerShell 5.1 or later.
- Windows operating system.
- Venmo transaction CSV file as input.

## Usage

1. Open a PowerShell prompt.
2. Run the script with the following command:

   ```powershell
   .\VenmoToQuicken.ps1 -InputCsv <path-to-venmo-csv> -OutputCsv <path-to-output-csv> -Account <account-name> -DateFormat <date-format>