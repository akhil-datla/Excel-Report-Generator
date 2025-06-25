/*
 * File: types.go
 * File Created: Thursday, 20th June 2024 11:15:02 am
 * Last Modified: Thursday, 20th June 2024 5:45:11 pm
 * Author: Akhil Datla
 * Copyright © Akhil Datla 2024
 */

package reportgen

import (
	"github.com/xuri/excelize/v2"
)

// SupportedDataType is an interface for different types of data that can be input into an Excel sheet.
type SupportedDataType interface {
	Size() int
	InputToSheet(excelFile *excelize.File, sheet string, cellRange string) error
}

// Scalar represents a single cell value.
type Scalar struct {
	Value any
}

// Size returns the size of the Scalar, which is always 1.
func (s Scalar) Size() int {
	return 1
}

// InputToSheet inputs the scalar value into the specified cell.
func (s Scalar) InputToSheet(excelFile *excelize.File, sheet string, cell string) error {
	return excelFile.SetCellValue(sheet, cell, s.Value)
}

// Vector represents a single column of values.
type Vector struct {
	Values []any
}

// Size returns the number of values in the Vector.
func (v Vector) Size() int {
	return len(v.Values)
}

// InputToSheet inputs the vector values into the specified cell range.
func (v Vector) InputToSheet(excelFile *excelize.File, sheet string, cellRange string) error {
	coordinates, err := rangeRefToCoordinates(cellRange)
	if err != nil {
		return err
	}
	if err := sortCoordinates(coordinates); err != nil {
		return err
	}
	for i, value := range v.Values {
		cellName, err := excelize.CoordinatesToCellName(coordinates[0], coordinates[1]+i)
		if err != nil {
			return err
		}
		if err := excelFile.SetCellValue(sheet, cellName, value); err != nil {
			return err
		}
	}
	return nil
}

// Table represents a two-dimensional array of values.
type Table struct {
	Values [][]any
}

// Size returns the total number of values in the Table.
func (t Table) Size() int {
	size := 0
	for _, row := range t.Values {
		size += len(row)
	}
	return size
}

// InputToSheet inputs the table values into the specified cell range.
func (t Table) InputToSheet(excelFile *excelize.File, sheet string, cellRange string) error {
	coordinates, err := rangeRefToCoordinates(cellRange)
	if err != nil {
		return err
	}
	if err := sortCoordinates(coordinates); err != nil {
		return err
	}
	for i, row := range t.Values {
		for j, value := range row {
			cellName, err := excelize.CoordinatesToCellName(coordinates[0]+j, coordinates[1]+i)
			if err != nil {
				return err
			}
			if err := excelFile.SetCellValue(sheet, cellName, value); err != nil {
				return err
			}
		}
	}
	return nil
}
