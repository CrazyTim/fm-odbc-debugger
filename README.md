<div>
  <img alt="thumbnail" src="https://crazytim.github.io/fm-odbc-debugger/repo-thumbnail.png"/>
  <br>
</div>

# FileMaker ODBC Debugger

FileMaker's ODBC driver has its own flavour of SQL syntax, and there are limitations and gotchas. This lightweight utility hides some of these limitations and allows you to use syntax more like T-SQL (see feature list below). Performance can also be an issue, and so the query execution time is split into three different metrics. Hopefully this makes your life easier when writing and testing queries.

## Disclosure

This program is provided "as is" and comes without a warranty of any kind. Use at your own risk!

## Getting Started

- Build the solution in Visual Studio.
- Install 64bit FileMaker ODBC Driver (not provided).

## Technical Notes

- Works with 64 bit drivers only.
- Warning: running concurrent queries to FileMaker can have a big impact on performance. Support for concurrent queries was sadly removed in v16, and since then all queries are executed one item at a time, significantly reducing the usefulness of their ODBC driver.
- The FileMaker ODBC Driver doesn't allow you to specify a port other than 2399.

## Features

- Show how long the query took to execute (separate durations for connect, process, and stream).
- Support for T-SQL style comments (single `--` and multi-line `/*`).
- Text editor supports syntax highlighting for all of the documented keywords. Colours are the same as Microsoft SQL Server Management Studio.
- Execute multiple queries as a transaction (separate queries using a semi-colon `;`).
- Convert line breaks to carriage returns (FileMaker natively uses `CR` for line breaks).
- Field names starting with an underscore will be automatically escaped with double quotes.
- Multiple tabs. Each tab has its own connection details. Tabs are remembered when opening the program again.
- Execute multiple queries at the same time (multi-threaded).
- Provide links to the official FileMaker references/guides.
- Right-click menu allows you to paste common FileMaker SQL functions.
- Show warnings for several syntax errors and gotchas, some of which are unique to FileMaker.
- Right-click on results and `Copy Selected Column(s) As CSV` (handy for executing a follow-up `WHERE IN (x,y,z)` query).
- Quickly search results for a keyword.
- Limit results to a maximum number of rows.
- Highlight NULL values yellow in the results.
- Show what version of the FileMaker ODBC driver is installed.
- Manually specify a driver name and connection string (you can use any installed 64bit ODBC driver).

## Acknowledgements:
- Icon by [David Vignoni](www.iconfinder.com/icons/1230/animal_bug_insect_ladybird_icon), GNU Lesser General Public License (LGPL).
- Tab control code adapted from [MDI tab control](https://www.codeproject.com/Articles/16436/A-highly-configurable-MDI-tab-control-from-scratch) by Eduardo Oliveira, The Code Project Open License (CPOL). 
- [CsvHelper](joshclose.github.io/CsvHelper), Apache License, Version 2.0.
- [Json.NET](https://www.newtonsoft.com/json), MIT license.
- [FastColoredTextBox](https://github.com/PavelTorgashov/FastColoredTextBox), GNU Lesser General Public License (LGPLv3).
