# TODO

## High

- Clean/replace any dodgy (non-UTF-8?) characters in query, such as `‘`. Example: `SELECT _jobid FROM contracts WHERE ( status = ‘onhire’ or status = ‘ready for pickup’) and product = ‘void'`
- Show a list of the ODBC drivers installed on the system.
- Get a list of the available tables.
- Better text editor:
    - support multiple undos (ctrl-z).
    - sql syntax highlighting.
    - capitalise SQL reserved words.

## Low

- Replace tab control with a better one:
  - smooth animations when rearranging tabs (convert project to wpf and implement [draggable tabs](https://dragablz.net)).
  - tabs disappear (and the order will rearrange) if one extends beyond the edge of the window.

- Ignore `[` and `]` characters in a query so we can use T-SQL field naming syntax.
  - refer: https://docs.microsoft.com/en-us/sql/relational-databases/databases/database-identifiers?view=sql-server-2017
  - refer: https://dba.stackexchange.com/a/7513/37320
