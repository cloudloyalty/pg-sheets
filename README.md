pg-sheets
=========

Export result of a Postgres SQL query to Google Spreadsheets.

Usage of `pg-sheets`:
```
  -query string
        SQL file to execute
  -dsn string
        database connection string
  -spreadsheet string
        spreadsheet ID string
  -sheet int
        sheet ID integer
  -append
        append to spreadsheet, not overwrite
  -header
        include header in result
  -credentials string
        credentials file (default "credentials.json")
  -token string
        token storage file (default "token.json")
```

Usage example:
```
./pg-sheets -query=query.sql \
            -dsn=postgresql://localhost/db?sslmode=disable \
            -spreadsheet=14KZv6vclr5CZ5IdTmnpc-XXYYYZZZ0001122233 \
            -sheet=0 \
            -header \
            -append
```

How to create `credentials.json`:
- https://developers.google.com/workspace/guides/create-credentials#desktop

At first run you will be prompted for the auth token that will be saved in `token.json` for later use.

How to obtain `spreadsheet` and `sheet` values from the URL:
```
https://docs.google.com/spreadsheets/d/14KZv6vclr5CZ5IdTmnpc-XXYYYZZZ0001122233/edit#gid=12345
                                       ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^          ^^^^^
                                                            `-- spreadsheet                `-- sheet
```
