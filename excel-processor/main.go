package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Function to process strings - customize this for your needs
func processString(input string) string {
	// This is a simple example - replace with your actual string processing logic
	return strings.ToUpper(input) + "_PROCESSED"
}

func processExcelFile(inputPath string, outputPath string, columnName string) error {
	// Open the Excel file
	f, err := excelize.OpenFile(inputPath)
	if err != nil {
		return fmt.Errorf("error opening Excel file: %w", err)
	}
	defer f.Close()

	// Get all sheet names
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return fmt.Errorf("no sheets found in Excel file")
	}

	// Process the first sheet
	sheetName := sheets[0]

	// Get all rows in the sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("error reading sheet: %w", err)
	}

	if len(rows) == 0 {
		return fmt.Errorf("empty sheet")
	}

	// Find the column index
	columnIndex := -1
	for i, cell := range rows[0] { // rows[0]为标题行
		if cell == columnName {
			columnIndex = i
			break
		}
	}

	if columnIndex == -1 {
		return fmt.Errorf("column %s not found", columnName)
	}

	// Process each row
	for i := 1; i < len(rows); i++ { // i为正文第一行索引

		// Skip rows that don't have enough columns
		if len(rows[i]) <= columnIndex {
			continue
		}

		// Get the cell value
		cellValue := rows[i][columnIndex]

		// Process the string
		processedValue := processString(cellValue)

		// Write the processed value to the next column
		rowNum := i + 1
		cellAddr, err := excelize.CoordinatesToCellName(columnIndex+2, rowNum) // +2 for next column (+1 would be the same, +2 is the right neighbor)
		if err != nil {
			return fmt.Errorf("error getting cell address: %w", err)
		}

		err = f.SetCellValue(sheetName, cellAddr, processedValue)
		if err != nil {
			return fmt.Errorf("error setting cell value: %w", err)
		}
	}

	// Save the Excel file
	err = f.SaveAs(outputPath)
	if err != nil {
		return fmt.Errorf("error saving file: %w", err)
	}

	return nil
}

func processCSVFile(inputPath string, outputPath string, columnIndex int) error {
	// Open the CSV file
	file, err := os.Open(inputPath)
	if err != nil {
		return fmt.Errorf("error opening CSV file: %w", err)
	}
	defer file.Close()

	// Create a CSV reader
	reader := csv.NewReader(file)

	// Read all records
	var records [][]string
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("error reading CSV: %w", err)
		}
		records = append(records, record)
	}

	if len(records) == 0 {
		return fmt.Errorf("empty CSV file")
	}

	// Verify column index exists in header
	if columnIndex < 0 || columnIndex >= len(records[0]) {
		return fmt.Errorf("column index %d out of bounds", columnIndex)
	}

	// Process each row (skip header)
	for i := 1; i < len(records); i++ {
		// Skip rows that don't have enough columns
		if len(records[i]) <= columnIndex {
			continue
		}

		// Get the cell value
		cellValue := records[i][columnIndex]

		// Process the string
		processedValue := processString(cellValue)

		// Add the processed value to the next column
		if columnIndex+1 >= len(records[i]) {
			// Extend the record if needed
			records[i] = append(records[i], processedValue)
		} else {
			// Insert into existing array
			records[i][columnIndex+1] = processedValue
		}
	}

	// Create output file
	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("error creating output file: %w", err)
	}
	defer outFile.Close()

	// Write to CSV
	writer := csv.NewWriter(outFile)
	defer writer.Flush()

	for _, record := range records {
		err := writer.Write(record)
		if err != nil {
			return fmt.Errorf("error writing to CSV: %w", err)
		}
	}

	return nil
}

func main() {
	// Define command-line flags
	inputFile := flag.String("input", "", "Path to the input Excel/CSV file")
	outputFile := flag.String("output", "", "Path to save the processed file")
	column := flag.String("column", "", "Column name for Excel or column index for CSV (0-based)")

	flag.Parse()

	// Validate input parameters
	if *inputFile == "" || *outputFile == "" || *column == "" {
		fmt.Println("All parameters are required")
		flag.PrintDefaults()
		return
	}

	// Determine file type
	ext := strings.ToLower(filepath.Ext(*inputFile))

	var err error
	switch ext {
	case ".xlsx":
		err = processExcelFile(*inputFile, *outputFile, *column)
	case ".csv":
		// For CSV files, convert column to index
		var columnIndex int
		_, err = fmt.Sscanf(*column, "%d", &columnIndex)
		if err != nil {
			fmt.Printf("Error: For CSV files, column must be a numeric index (0-based): %v\n", err)
			return
		}
		err = processCSVFile(*inputFile, *outputFile, columnIndex)
	default:
		fmt.Printf("Unsupported file format: %s\n", ext)
		return
	}

	if err != nil {
		fmt.Printf("Error processing file: %v\n", err)
		return
	}

	fmt.Println("File processed successfully!")
}
