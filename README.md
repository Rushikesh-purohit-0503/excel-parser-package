# excelparser

A small Go package to parse Excel (.xlsx) files into JSON-friendly structures with header auto-detection, header filtering/mapping, trimming, and optional concurrency when parsing multiple sheets.

## Features

- Detects header row automatically (configurable scan limit).
- Map header names to custom keys.
- Filter which headers to include.
- Trim cell whitespace and optionally skip empty values.
- Parse multiple sheets concurrently with a configurable concurrency limit.

## Install

From your project directory run:

```bash
go get github.com/Rushikesh-purohit-0503/excel-parser-package
```

Or add the package path to your `go.mod` and run `go mod tidy`.

The package depends on `github.com/xuri/excelize/v2` which will be fetched automatically by Go modules.

## Usage

Here's a minimal example using the repository's `main.go` pattern:

```go
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
		SheetNames:          []string{}, // empty = parse all sheets
		HeaderRowAutoDetect: true,
		HeaderRowScanLimit:  50,
		HeaderMap:           map[string]string{},
		TrimSpace:           true,
		SkipEmpty:           true,
		MaxConcurrentSheets: 5,
	}

	result, err := excelparser.ParseExcel("test.xlsx", opts)
	if err != nil {
		log.Fatal(err)
	}

	b, _ := json.MarshalIndent(result, "", "  ")
	_ = os.WriteFile("output.json", b, 0644)
	fmt.Println("✅ JSON file created: output.json")
}
```

## ParseOptions

- `SheetNames []string` — list of sheet names to parse. If empty, all sheets will be parsed.
- `HeaderFilter []string` — if non-empty, only headers matching (case-insensitive) items in this list are included.
- `HeaderMap map[string]string` — map from original header name -> new key to use in output records.
- `TrimSpace bool` — trim leading/trailing spaces from header and cell values when true.
- `SkipEmpty bool` — when true, skips empty cell values when building a record.
- `HeaderRowAutoDetect bool` — if true, the parser will attempt to find the header row automatically.
- `HeaderRowScanLimit int` — how many rows to scan when auto-detecting the header (0 or negative means default).
- `MaxConcurrentSheets int` — maximum number of sheets to parse concurrently; 0 means unlimited (uses Goroutines without throttling).

## Output (SheetResult)

Each parsed sheet is returned as a `SheetResult` keyed by sheet name:

- `Headers []string` — the header row (as read from the sheet, before header mapping).
- `Records []map[string]string` — slice of records (each record is a map of header->value or mapped header->value).
- `RecordCount int` — number of records captured for the sheet.

Example `output.json` snippet:

```json
{
	"Sheet1": {
		"headers": ["Name", "Date", "Amount"],
		"records": [
			{"Name": "Alice", "Date": "2023-01-01", "Amount": "100"}
		],
		"recordCount": 1
	}
}
```

## Notes & Tips

- The header auto-detection attempts to find a row with at least 3 non-empty cells followed by multiple subsequent rows with data — this works well for typical tabular sheets but may need adjustment for exotic layouts.
- `HeaderMap` only applies after a header row is selected; the original header text is used to look up mapping keys.
- Use `MaxConcurrentSheets` to limit memory/CPU when working with very large workbooks.

## Contributing

Contributions and bug reports welcome. Please open an issue or a pull request in the repository.

## License
This project is licensed under the MIT License — see the `LICENSE` file in this repository for the full text.

Copyright (c) 2025 Rushikesh Purohit