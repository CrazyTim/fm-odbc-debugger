<div>
  <img alt="thumbnail" src="https://crazytim.github.io/fm-odbc-debugger/repo-thumbnail.png"/>
  <br>
</div>

# FileMaker ODBC Debugger

FileMaker's ODBC driver has its own flavour of SQL syntax, and there are limitations and gotchas. This lightweight utility hides some of these limitations and allows you to use syntax more like T-SQL. Performance can also be an issue, and so the query execution time is split into three different metrics. All this and more hopefully makes your life easier when working with FileMaker over ODBC.

## Disclosure

This program is provided "as is" and comes without a warranty of any kind. Use at your own risk!

## Getting Started

Build the solution in Visual Studio.

## Technical Notes

- Works with 64 bit drivers only.
- Warning: running concurrent queries to FileMaker can have a big impact on performance. Support for concurrent queries was sadly removed in v16, and since then all queries are executed one item at a time, significantly reducing the usefulness of their ODBC driver.
- The FileMaker ODBC Driver doesn't allow you to specify a port other than 2399.

## Features

- Show how long the query took to run (separate connect, process, and stream times).
- Support for T-SQL single and multi-line comments.
- Execute multiple queries as a transaction (separate queries using a semi-colon (;) like T-SQL).
- Convert line breaks from Windows (CRLF) to FileMaker (CR).
- Automatically surround field names starting with an underscore with double quotes.
- Show the version of the FileMaker ODBC driver.
- Multiple tabs. Each tab has its own connection details and remembers what is on each tab after closing down.
- Multi-threaded - execute multiple queries at the same time.
- Links to official FileMaker reference guides.
- Right-click to insert common FileMaker functions.
- Show warnings for several syntax errors and gotchas, some of which are unique to FileMaker.
- Right-click on results and "copy column as CSV" - handy for executing a follow-up "WHERE IN (x,y,z)" query.
- Manually specify a driver name and connection string (you can use any installed ODBC driver).
- Quickly search for a keyword in the results panel.
- Limit results to a maximum number of rows.
- Highlight NULL cells yellow.

## Acknowledgements:
- Icon by David Vignoni (www.iconfinder.com/icons/1230/animal_bug_insect_ladybird_icon).
- Csv code from joshclose.github.io/CsvHelper, Apache License, Version 2.0.
- Tab control code adapted from unknown.
