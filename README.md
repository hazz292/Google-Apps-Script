# Google-Apps-Script
Example of Google Apps Script used to automate reporting.
Code is organised into 3 files:
1. shared.gs - shared functions
2. functions.gs - modular approach to perform each action
3. createReport.gs - calls difference functions to generate report

General logic as follows:
1. Pull in data from Metabase page using public url
2. Populate data in spreadsheet as data source
3. Pull into google slide template for reporting
