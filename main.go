package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"

	"github.com/Rushikesh-purohit-0503/excelparser"
)

func main() {
	opts := excelparser.ParseOptions{
		SheetNames:          []string{},
		HeaderRowAutoDetect: true,
		HeaderRowScanLimit:  50,
		// HeaderFilter:        []string{"Name", "Amount", "Date"},
		HeaderMap:           map[string]string{},
		TrimSpace:           true,
		SkipEmpty:           true,
		MaxConcurrentSheets: 5, // Limit to 2 concurrent sheet parsing
	}
	result, err := excelparser.ParseExcel("test.xlsx", opts)
	if err != nil {
		log.Fatal(err)
	}
	jsonData, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		panic(err)
	}
	err = os.WriteFile("output.json", jsonData, 0644)
	if err != nil {
		panic(err)
	}

	fmt.Println("✅ JSON file created: output.json")

}
