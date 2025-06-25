/*
 * File: reportgen.go
 * File Created: Thursday, 20th June 2024 11:14:00 am
 * Last Modified: Friday, 21st June 2024 10:25:05 am
 * Author: Akhil Datla
 * Copyright © Akhil Datla 2024
 */

package reportgen

import (
	"bytes"

	"github.com/xuri/excelize/v2"
)

// Options holds configuration options for generating a report.
type Options struct {
	TemplateFilePath string
	OutputFilePath   string
	NewFile          bool
	ExcelizeOpts     excelize.Options
}

// ReportGenerator is responsible for generating reports using an existing Excel file as a template.
type ReportGenerator struct {
	ExcelFile *excelize.File
	Opts      Options
}

// NewReportGenerator creates a new ReportGenerator instance using the provided options.
func NewReportGenerator(opts Options) (*ReportGenerator, error) {
	if opts.NewFile {
		return &ReportGenerator{
			ExcelFile: excelize.NewFile(),
			Opts:      opts,
		}, nil
	}
	f, err := excelize.OpenFile(opts.TemplateFilePath, opts.ExcelizeOpts)
	if err != nil {
		return nil, err
	}
	return &ReportGenerator{
		ExcelFile: f,
		Opts:      opts,
	}, nil
}

// InputData inputs data into the specified sheet and cell range of the Excel file.
func (rg *ReportGenerator) InputData(sheet string, cellRange string, data SupportedDataType) error {
	switch data.(type) {
	case Scalar:
		return data.InputToSheet(rg.ExcelFile, sheet, cellRange)
	case Vector:
		coordinates, err := rangeRefToCoordinates(cellRange)
		if err != nil {
			return err
		}
		if err := sortCoordinates(coordinates); err != nil {
			return err
		}

		if err := checkSize(data, coordinates); err != nil {
			return err
		}

		return data.InputToSheet(rg.ExcelFile, sheet, cellRange)
	case Table:
		coordinates, err := rangeRefToCoordinates(cellRange)
		if err != nil {
			return err
		}
		if err := sortCoordinates(coordinates); err != nil {
			return err
		}

		if err := checkSize(data, coordinates); err != nil {
			return err
		}

		return data.InputToSheet(rg.ExcelFile, sheet, cellRange)
	default:
		return ErrDataTypeInvalid
	}
}

// SaveSheet saves the Excel file, either overwriting the existing file or creating a new one at OutputFilePath based on the NewFile option.
func (rg *ReportGenerator) SaveSheet() error {
	if rg.Opts.NewFile {
		return rg.ExcelFile.SaveAs(rg.Opts.OutputFilePath, rg.Opts.ExcelizeOpts)
	}
	return rg.ExcelFile.Save()
}

// GetFileBuffer returns the Excel file buffer.
func (rg *ReportGenerator) GetFileBuffer() (*bytes.Buffer, error) {
	return rg.ExcelFile.WriteToBuffer()
}
