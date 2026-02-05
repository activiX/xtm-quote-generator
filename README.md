# XTM Quote Input Generator

Console application written in C# that processes single or multiple XTM Excel analysis files
and generates a consolidated quote spreadsheet.

The tool reads word count data from a fixed XTM analysis template and outputs a
single Excel file ready to be used for quote preparation.

## Features
- Processes single or multiple XTM Excel analysis files in batch
- Validates template structure before processing
- Aggregates word counts across match categories
- Maps language codes using a CSV mapping file so new languages can be added by any team member without code changes
- Generates a sortable Excel output using ClosedXML

## Purpose
- Speeds up quote creation for clients, which was previously done manually file by file

## Files
- `Program.cs` – main application logic
- `language-map.csv` – mapping of language codes to display names used in the output quote

## Requirements
- .NET 9.0
- ClosedXML

## How to run
Run the application as a standard C# console application.
