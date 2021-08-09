pg-sheets
=========

Export result of a Postgres SQL query to Google Spreadsheets.

Usage:
```
./pg-sheets -query=query.sql \
            -dsn=postgresql://localhost/db?sslmode=disable \
            -spreadsheet=14KZv6vclr5CZ5IdTmnpc-XXYYYZZZ0001122233 \
            -sheet=0 \
            -append
```
