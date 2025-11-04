package excelparser

import (
	"fmt"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

type ParseOptions struct {
	SheetNames          []string
	HeaderFilter        []string
	HeaderMap           map[string]string
	TrimSpace           bool
	SkipEmpty           bool
	HeaderRowAutoDetect bool
	HeaderRowScanLimit  int
	MaxConcurrentSheets int
}

type SheetResult struct {
	Headers     []string            `json:"headers"`
	Records     []map[string]string `json:"records"`
	RecordCount int                 `json:"recordCount"`
}

func ParseExcel(filePath string, opts ParseOptions) (map[string]SheetResult, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open excel file: %w", err)
	}
	defer f.Close()

	result := make(map[string]SheetResult)
	sheets := opts.SheetNames
	if len(sheets) == 0 {
		sheets = f.GetSheetList()
	}

	var (
		wg   sync.WaitGroup
		mu   sync.Mutex
		sema chan struct{}
	)

	if opts.MaxConcurrentSheets > 0 {
		sema = make(chan struct{}, opts.MaxConcurrentSheets)
	}

	for _, sheet := range sheets {
		wg.Add(1)
		sheetName := sheet

		go func() {
			defer wg.Done()

			if sema != nil {
				sema <- struct{}{}
				defer func() { <-sema }()
			}

			rows, err := f.GetRows(sheetName)
			if err != nil {
				fmt.Printf("⚠️ failed to read sheet %s: %v\n", sheetName, err)
				return
			}
			if len(rows) == 0 {
				return
			}

			// Detect header row
			headerRowIdx := 0
			if opts.HeaderRowAutoDetect {
				limit := opts.HeaderRowScanLimit
				if limit <= 0 {
					limit = 10
				}
				headerRowIdx = detectHeaderRow(rows, limit)
			}

			headers := rows[headerRowIdx]
			headerIndexes := map[int]string{}
			for idx, h := range headers {
				if opts.TrimSpace {
					h = strings.TrimSpace(h)
				}
				if len(opts.HeaderFilter) == 0 || containsIgnoreCase(opts.HeaderFilter, h) {
					headerIndexes[idx] = h
				}
			}

			// Parse rows
			localData := []map[string]string{}
			for _, row := range rows[headerRowIdx+1:] {
				record := map[string]string{}
				for idx, header := range headerIndexes {
					if idx < len(row) {
						val := row[idx]
						if opts.TrimSpace {
							val = strings.TrimSpace(val)
						}
						if opts.SkipEmpty && val == "" {
							continue
						}

						if newName, ok := opts.HeaderMap[header]; ok {
							record[newName] = val
						} else {
							record[header] = val
						}
					}
				}
				if len(record) > 0 {
					localData = append(localData, record)
				}
			}

			sheetResult := SheetResult{
				Headers:     headers,
				Records:     localData,
				RecordCount: len(localData),
			}

			mu.Lock()
			result[sheetName] = sheetResult
			mu.Unlock()
		}()
	}

	wg.Wait()
	return result, nil
}

func detectHeaderRow(rows [][]string, scanLimit int) int {
	if len(rows) == 0 {
		return 0
	}

	if scanLimit <= 0 || scanLimit > len(rows) {
		scanLimit = len(rows)
	}

	for i := 0; i < scanLimit; i++ {
		headerRow := rows[i]
		nonEmptyHeader := 0
		for _, cell := range headerRow {
			if strings.TrimSpace(cell) != "" {
				nonEmptyHeader++
			}
		}

		// Skip rows that are too empty
		if nonEmptyHeader < 3 {
			continue
		}

		// Check next 5 rows — should have data
		validDataRows := 0
		for j := i + 1; j < len(rows) && j < i+6; j++ { // look ahead 5 rows
			dataNonEmpty := 0
			for _, val := range rows[j] {
				if strings.TrimSpace(val) != "" {
					dataNonEmpty++
				}
			}
			if dataNonEmpty >= 3 {
				validDataRows++
			}
		}

		// If at least 4–5 rows below are non-empty → real header
		if validDataRows >= 4 {
			return i
		}
	}

	// fallback
	return 0
}

func containsIgnoreCase(arr []string, target string) bool {
	for _, a := range arr {
		if strings.EqualFold(a, target) {
			return true
		}
	}
	return false
}
