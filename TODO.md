# TODO

## High

- FileMaker's ODBC Driver only supports UTF-8. Clean/replace any dodgy (non-UTF-8?) characters in query, such as `‘`. Example: `SELECT id FROM test WHERE test = ‘one’ OR test = ‘two'`. Alternatively show an error about the presence of non-UTF-8 characters in the query.
- If the driver reported the line and character where it found a syntax issue, display this in a more user friendly way. Example: `ERROR: FQL001/(2:4): There is an error in the syntax of the query`, means line 2, character 4.
- Ignore `[` and `]` characters around field names so we can accept T-SQL field naming syntax if the query was pasted from SQLSMS.
  - refer: https://docs.microsoft.com/en-us/sql/relational-databases/databases/database-identifiers?view=sql-server-2017
  - refer: https://dba.stackexchange.com/a/7513/37320

## Low

- Show a list of the ODBC drivers installed on the system.
- Get a list of the available tables in the FileMaker database.
- Replace tab control with a better one:
  - smooth animations when rearranging tabs (convert project to wpf and implement [draggable tabs](https://dragablz.net)).
  - tabs disappear (and the order will rearrange) if one extends beyond the edge of the window.
