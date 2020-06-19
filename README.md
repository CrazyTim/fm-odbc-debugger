<div>
  <img alt="thumbnail" src="https://crazytim.github.io/fm-odbc-debugger/repo-thumbnail.png"/>
  <br>
</div>

# FileMaker ODBC Debugger

Write and test SQL queries for a FileMaker database.

## Motivation

FileMaker's ODBC driver has its own flavour of SQL syntax, as well as some lesser-known quirks. This tool helps you write and test SQL queries to a FileMaker database while handling some of these quirks. It also warns you about things to avoid, and supports syntax highlighting for all of FileMakers' reserved keywords (see [full feature list below](#features)).

In addition, FileMaker ODBC performance can be very slow, which can often be solved by optimising the query. So to help with performance testing it shows the total execution time split into 3 different times; 'connect', 'execute', and 'stream'. This can sometimes give insight into where the bottleneck is (calculation fields are usually the culprit).

To dig deeper with optimising FileMaker queries, refer to this [article about the `ExecuteSQL` function](https://www.soliantconsulting.com/blog/executesql-filemaker-performance/).

## Installing

First get familiar with the [FileMaker ODBC Guide](https://fmhelp.filemaker.com/docs/16/en/fm16_odbc_jdbc_guide.pdf) and follow [these instructions](https://fmhelp.filemaker.com/help/16/fmp/en/#page/FMP_Help%2Fsharing-via-odbc-jdbc.html%23) for sharing your FileMaker database via ODBC.

- Install the 64bit FileMaker ODBC driver that comes bundled with FileMaker.
- Download the latest release of [FileMaker ODBC Debugger](https://github.com/CrazyTim/fm-odbc-debugger/releases).
- Launch FileMaker ODBC Debugger and ensure the correct `Driver` version has been detected. 
- Enter values for `Server` (name or ip address), `Database` (name), and `Credentials` (username/password). Type your query and then execute it!

## Technical Notes

- Works with 64 bit drivers only.
- Depending on the data, concurrent queries to FileMaker can be very slow. With some experimentation, tweaking, and simplifying the query, you might be able to get it to acceptable speeds. Sadly, support for concurrent queries was [removed in v16 onwards](https://community.claris.com/en/s/question/0D50H00006ezLy6/issue-with-concurrent-odbc-connections-in-fms16-and-fms17), significantly reducing the usefulness of FileMaker's ODBC driver.
- The FileMaker ODBC Driver doesn't allow you to specify a port other than 2399.

## Features

- Allow comments, single-line `--` and multi-line `/*`.
- Execute multiple statements as a transaction, separated by a semi-colon `;`.
- Convert all line breaks to carriage returns (FileMaker uses `CR` for line breaks).
- Field names starting with an underscore will be automatically escaped with double quotes.
- Show how long the query took to execute (separate durations for connect, process, and stream).
- Text editor supports syntax highlighting for all of the supported FileMaker SQL keywords (There are reserved keywords that can clash with your field and table names). Colours are similar to Microsoft SQL Server Management Studio.
- Multiple tabs. Each tab has its own connection details. Tabs are remembered when opening the program again.
- Execute two or more queries in different tabs at the same time.
- Show warnings for several FileMaker-specific syntax issues:
    - FileMaker stores empty strings as NULL, so using `WHERE column = ''` or `WHERE column <> ''` on a `Text` field will always return 0 results. Use `IS NULL` instead.
    - FileMaker ODBC does not support the keywords `TRUE` or `FALSE`. Use `1` or `0` instead.
    - Query contains the keyword `BETWEEN` and FileMaker ODBC is VERY slow when comparing dates this way. Use `>=` and `<=` instead.
- Provide links to the official FileMaker references/guides (in right-click menu).
- Right-click menu allows you to paste common FileMaker SQL functions.
- Right-click on query results and `Copy Selected Column(s) As CSV` (handy for executing a follow-up `WHERE IN (x,y,z)` query).
- Quickly search query results for a keyword.
- Limit query results to a maximum number of rows.
- Highlight NULL values yellow in the query results.
- Show what version of the FileMaker ODBC driver is installed.
- Option to manually specify a driver name and connection string (you can use any installed 64bit ODBC driver).

## Compiling

- You will need to install the [Wix Toolset Visual Studio 2019 Extension](https://marketplace.visualstudio.com/items?itemName=WixToolset.WixToolsetVisualStudio2019Extension).

## Acknowledgements
- Icon by [David Vignoni](www.iconfinder.com/icons/1230/animal_bug_insect_ladybird_icon), GNU Lesser General Public License (LGPL).
- Tab control code adapted from [MDI tab control](https://www.codeproject.com/Articles/16436/A-highly-configurable-MDI-tab-control-from-scratch) by Eduardo Oliveira, The Code Project Open License (CPOL). 
- [CsvHelper](joshclose.github.io/CsvHelper), Apache License, Version 2.0.
- [Json.NET](https://www.newtonsoft.com/json), MIT license.
- [FastColoredTextBox](https://github.com/PavelTorgashov/FastColoredTextBox), GNU Lesser General Public License (LGPLv3).
