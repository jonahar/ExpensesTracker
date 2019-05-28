# ExpensesTracker

This project consists of a javascript code to be integrated into Google sheets (via
Google Apps Script) and build a customized tool to track and get statistics of 
one's monthly expenses.

The spreadsheet should have 2 sheets named _Expenses_ and _Summary_.
The header and line format in the _Expenses_ sheet should be as described below:


| Date       | Description    | Amount    | Category    |
|------------|----------------|-----------|-------------|
| YYYY-MM-DD | \<description> | \<number> | \<category> |

As the above table suggests, a date format must be 'YYYY-MM-DD'. The _Summary_ sheet should not contain user-entered information, and is filled automatically.

The provided code adds 2 new buttons to google sheets toolbar, that compute statistics and display a nice chart.

