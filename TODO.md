# TODO

## High

- Clean/replace any dodgy (non-UTF-8?) characters in query, such as `‘`.
  - example: `SELECT _jobid FROM contracts WHERE ( status = ‘onhire’ or status = ‘ready for pickup’) and product = ‘void'`

- Better text editor:
    - support multiple undos.
    - sql syntax highlighting.

## Low

- Replace tab control with a better one:
  - smooth animations when rearranging tabs (convert project to wpf and implement [draggable tabs](https://dragablz.net)).
  - tabs just disappear if not enough room.

- Ignore `[` and `]` characters in a query so we can use T-SQL field naming syntax.
  - refer: https://docs.microsoft.com/en-us/sql/relational-databases/databases/database-identifiers?view=sql-server-2017
  - refer: https://dba.stackexchange.com/a/7513/37320
