# Json-to-excel-updator

This Python utility reads multiple JSON files containing employee utilisation data, identifies the correct rows in corresponding Excel files based on `empid` and `date`, and updates only the relevant cells in-place. It also creates backups of Excel files and archives processed JSON files.

## Features

- Updates only specific cells in Excel (`.xlsx`) files based on `empid` and `date`.
- Reads and processes multiple JSON files from a folder.
- Supports multiple Excel files (one per team).
- Backs up Excel files before updating.
- Archives JSON files after processing.
- Preserves all existing Excel formatting and formulas.

## Folder Structure

