package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/joho/godotenv"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type (
	VendorReport struct {
		Date       time.Time
		Env        string
		Vendor     string
		Action     string
		URL        string
		Request    string
		Response   string
		HttpStatus int
		Elapsed    int
		Retry      int
		Error      string
	}
)

func getSheetConfig() *sheets.Service {
	b, err := os.ReadFile("credentials-account.json")
	if err != nil {
		log.Fatalf("Failed to get key: %v", err)
	}

	srv, err := sheets.NewService(context.Background(), option.WithCredentialsJSON(b))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	return srv
}

func AppendRow(spreadsheetId, sheetName string, vendorReport VendorReport, srv *sheets.Service) (*sheets.AppendValuesResponse, error) {
	// Get Latest Index
	valueRange, err := GetAllValueRange(spreadsheetId, sheetName, srv)
	if err != nil {
		log.Fatalf("Unable to get last row: %v", err)
	}
	idx := len(valueRange.Values)

	// Insert Values
	values := &sheets.ValueRange{
		Values: [][]interface{}{{
			idx, // No (A)
			vendorReport.Date.Format("2006-01-02 15:04:05"), // Date (B)
			vendorReport.Env,        // Env (C)
			vendorReport.Vendor,     // Vendor (D)
			vendorReport.Action,     // Action (E)
			vendorReport.URL,        // URL (F)
			vendorReport.Request,    // Request (G)
			vendorReport.Response,   // Response (H)
			vendorReport.HttpStatus, // HttpStatus (I)
			vendorReport.Elapsed,    // Elapsed (J)
			vendorReport.Retry,      // Retry (K)
			vendorReport.Error,      // Error (L)
		}},
	}
	return srv.Spreadsheets.Values.Append(spreadsheetId, sheetName+"!A:K", values).ValueInputOption("USER_ENTERED").Do()
}

// GetAllValueRange include header
func GetAllValueRange(spreadsheetId, sheetName string, srv *sheets.Service) (*sheets.ValueRange, error) {
	return srv.Spreadsheets.Values.Get(spreadsheetId, sheetName+"!A1:J").Do()
}

func Update(spreadsheetId, sheetName string, srv *sheets.Service) {

	values := &sheets.ValueRange{
		Values: [][]interface{}{{
			"Japan",
			"Software Engineer Lead",
		}},
	}

	_, err := srv.Spreadsheets.Values.Update(spreadsheetId, sheetName+"!B2:C2", values).ValueInputOption("USER_ENTERED").Do()

	if err != nil {
		log.Fatalf("Unable to insert data to sheet: %v", err)
	}

}

func ClearRow(spreadsheetId, sheetName string, srv *sheets.Service) {

	cvr := &sheets.ClearValuesRequest{}
	_, err := srv.Spreadsheets.Values.Clear(spreadsheetId, sheetName+"!A7:E7", cvr).Do()

	if err != nil {
		log.Fatalf("Unable to clear data from sheet: %v", err)
	}

}

func ClearCell(spreadsheetId, sheetName string, srv *sheets.Service) {

	cvr := &sheets.ClearValuesRequest{}
	_, err := srv.Spreadsheets.Values.Clear(spreadsheetId, sheetName+"!B2", cvr).Do()

	if err != nil {
		log.Fatalf("Unable to clear data from sheet: %v", err)
	}

}

func ClearColumn(spreadsheetId, sheetName string, srv *sheets.Service) {

	cvr := sheets.ClearValuesRequest{}
	_, err := srv.Spreadsheets.Values.Clear(spreadsheetId, sheetName+"!C2:C", &cvr).Do()

	if err != nil {
		log.Fatalf("Unable to clear data from sheet: %v", err)
	}

}

func GetCellValue(spreadsheetId, sheetName string, srv *sheets.Service) {
	values, err := srv.Spreadsheets.Values.Get(spreadsheetId, sheetName+"!A2:E7").Do()

	if err != nil {
		log.Fatalf("Unable to Get data from sheet: %v", err)
	}

	for _, value := range values.Values {
		fmt.Println(value)
	}
}

func main() {

	// Load .env to environment
	err := godotenv.Load(".env")
	if err != nil {
		log.Fatalf("Error loading .env file")
	}

	spreadsheetId := os.Getenv("SPREADSHEET_ID")
	sheetName := os.Getenv("SPREADSHEET_NAME")

	srv := getSheetConfig()

	// APPEND ROW
	vendorReport := VendorReport{
		Date:       time.Now(),
		Env:        "prod",                               // Env
		Vendor:     "sicepat",                            // Vendor
		Action:     "pickup",                             // Action
		URL:        "/pickup",                            // URL
		Request:    `{"referenceNumber" : "0010101010"}`, // Request
		Response:   `{"status" : "200"}`,                 // Response
		HttpStatus: 200,                                  // HttpStatus
		Elapsed:    500,                                  // Elapsed
		Retry:      2,                                    // Retry
		Error:      "error json.Unmarshall nil value",    // Error
	}
	appendRes, err := AppendRow(spreadsheetId, sheetName, vendorReport, srv)
	if err != nil {
		log.Fatalf("Unable to insert data to sheet: %v", err)
	}
	b, err := appendRes.MarshalJSON()
	if err != nil {
		log.Fatalf("Unable to insert data to sheet: %v", err)
	}
	println("[go-sheet] result:", string(b))

	// TOTAL DATA
	valueRange, err := GetAllValueRange(spreadsheetId, sheetName, srv)
	if err != nil {
		log.Fatalf("Unable to get last row: %v", err)
	}
	println("[go-sheet] totalData:", len(valueRange.Values)-1)
}
