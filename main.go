package main

import (
	"context"
	"database/sql"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"time"

	_ "github.com/lib/pq"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	sheets "google.golang.org/api/sheets/v4"
)

var flags struct {
	DSN             string
	QueryFile       string
	SpreadsheetID   string
	SheetID         int64
	CellRange       string
	Append          bool
	CredentialsFile string
	TokenFile       string
	IncludeHeader   bool
}

func main() {
	flag.StringVar(&flags.DSN, "dsn", "", "database connection string")
	flag.StringVar(&flags.QueryFile, "query", "", "SQL file to execute")
	flag.StringVar(&flags.SpreadsheetID, "spreadsheet", "", "spreadsheet ID string")
	flag.Int64Var(&flags.SheetID, "sheet", 0, "sheet ID integer")
	flag.BoolVar(&flags.Append, "append", false, "append to spreadsheet, not overwrite")
	flag.StringVar(&flags.CredentialsFile, "credentials", "credentials.json", "credentials file")
	flag.StringVar(&flags.TokenFile, "token", "token.json", "token storage file")
	flag.BoolVar(&flags.IncludeHeader, "header", false, "include header in result")

	flag.Parse()

	fmt.Printf("flags=%v\n", flags)

	dbh, err := sql.Open("postgres", flags.DSN)
	if err != nil {
		log.Fatalf("Unable to open database connection: %v", err)
	}

	q, err := ioutil.ReadFile(flags.QueryFile)
	if err != nil {
		log.Fatalf("Unable to read query file: %v", err)
	}

	b, err := ioutil.ReadFile(flags.CredentialsFile)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config, flags.TokenFile)

	ctx := context.Background()
	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	fmt.Println("Executing query...")

	rows, err := dbh.QueryContext(ctx, string(q))
	if err != nil {
		log.Fatalf("Failed to execute query: %v", err)
	}

	cols, err := rows.Columns()
	if err != nil {
		log.Fatalf("Failed to detect result columns: %v", err)
	}

	if len(cols) == 0 {
		log.Fatalf("No columns in result.")
	}

	var resultRows []*sheets.RowData

	if flags.IncludeHeader && !flags.Append {
		cells := make([]*sheets.CellData, len(cols))
		for i, v := range cols {
			cells[i] = makeCell(v)
		}
		resultRows = append(resultRows, &sheets.RowData{Values: cells})
	}

	dataRow := make([]interface{}, len(cols))
	dataRowRefs := make([]interface{}, len(cols))

	for i := range cols {
		dataRowRefs[i] = &dataRow[i]
	}

	for rows.Next() {
		err = rows.Scan(dataRowRefs...)
		if err != nil {
			log.Fatalf("Failed to scan row: %v", err)
		}
		resultRows = append(resultRows, makeRow(dataRow))
	}

	fmt.Println("Updating spreadsheet... ")

	var reqs []*sheets.Request

	if flags.Append {
		reqs = append(reqs, &sheets.Request{
			AppendCells: &sheets.AppendCellsRequest{
				Fields:  "userEnteredValue",
				SheetId: flags.SheetID,
				Rows:    resultRows,
			},
		})
	} else {
		reqs = append(reqs, &sheets.Request{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Fields: "gridProperties(rowCount,columnCount)",
				Properties: &sheets.SheetProperties{
					SheetId: flags.SheetID,
					GridProperties: &sheets.GridProperties{
						RowCount:    int64(len(resultRows)),
						ColumnCount: int64(len(cols)),
					},
				},
			},
		})
		reqs = append(reqs, &sheets.Request{
			UpdateCells: &sheets.UpdateCellsRequest{
				Fields: "userEnteredValue",
				Start: &sheets.GridCoordinate{
					SheetId:     flags.SheetID,
					RowIndex:    0,
					ColumnIndex: 0,
				},
				Rows: resultRows,
			},
		})
	}

	call := srv.Spreadsheets.BatchUpdate(
		flags.SpreadsheetID,
		&sheets.BatchUpdateSpreadsheetRequest{
			Requests: reqs,
		},
	)
	_, err = call.Context(ctx).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}

	fmt.Println("Successfully written.")
}

func makeRow(r []interface{}) *sheets.RowData {
	cells := make([]*sheets.CellData, len(r))
	for i, v := range r {
		cells[i] = makeCell(v)
	}
	return &sheets.RowData{Values: cells}
}

func makeCell(v interface{}) *sheets.CellData {
	toNumberValue := func(f float64) *sheets.CellData {
		return &sheets.CellData{
			UserEnteredValue: &sheets.ExtendedValue{
				NumberValue: &f,
			},
		}
	}
	toStringValue := func(s string) *sheets.CellData {
		return &sheets.CellData{
			UserEnteredValue: &sheets.ExtendedValue{
				StringValue: &s,
			},
		}
	}
	toBoolValue := func(b bool) *sheets.CellData {
		return &sheets.CellData{
			UserEnteredValue: &sheets.ExtendedValue{
				BoolValue: &b,
			},
		}
	}

	switch v := v.(type) {
	case string:
		return toStringValue(v)
	case int:
		return toNumberValue(float64(v))
	case int64:
		return toNumberValue(float64(v))
	case float64:
		return toNumberValue(v)
	case bool:
		return toBoolValue(v)
	case time.Time:
		// convert datetime to Excel "serial number"
		base := time.Date(2000, 1, 1, 0, 0, 0, 0, v.Location())
		days := float64(v.Unix()-base.Unix()) / 86400
		return toNumberValue(days + 36526) // 36526 is the serial number for 01.01.2000
	default:
		return toStringValue(fmt.Sprintf("unparsed: %T", v))
	}
}
