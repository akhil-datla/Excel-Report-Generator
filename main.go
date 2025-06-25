/*
 * File: main.go
 * File Created: Thursday, 20th June 2024 1:28:55 pm
 * Last Modified: Thursday, 20th June 2024 5:48:35 pm
 * Author: Akhil Datla
 * Copyright © Akhil Datla 2024
 */

package main

import (
	"fmt"
	"main/reportgen"
)

func main() {
	// Define options
	opts := reportgen.Options{
		TemplateFilePath: "template.xlsx",
		OutputFilePath:   "output.xlsx",
		NewFile:          true,
	}

	// Create a new ReportGenerator instance
	reportGenerator, err := reportgen.NewReportGenerator(opts)
	if err != nil {
		panic(err)
	}

	// Input scalar data into the sheet
	scalar := reportgen.Scalar{Value: "hello"}
	err = reportGenerator.InputData("Sheet1", "A1", scalar)
	if err != nil {
		panic(err)
	}

	// Input vector data into the sheet
	vector := reportgen.Vector{Values: []interface{}{1, 2, 3, 4, 5}}
	err = reportGenerator.InputData("Sheet1", "B1:B5", vector)
	if err != nil {
		panic(err)
	}

	// Input table data into the sheet
	table := reportgen.Table{
		Values: [][]interface{}{
			{1, 2, 3},
			{4, 5, 6},
			{7, 8, 9},
		},
	}
	err = reportGenerator.InputData("Sheet1", "C1:E3", table)
	if err != nil {
		panic(err)
	}

	// Save the sheet
	err = reportGenerator.SaveSheet()
	if err != nil {
		panic(err)
	}

	fmt.Println("Report generated successfully")
}
