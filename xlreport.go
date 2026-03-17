// Package xlreport provides a fluent API for generating Excel reports from Go.
//
// It wraps the excelize library with an ergonomic, chainable interface that makes
// it simple to create professional Excel reports with data, formulas, and styling.
//
// Basic usage:
//
//	r := xlreport.New()
//	defer r.Close()
//
//	r.Sheet("Sales").
//		Cell("A1", "Product").
//		Cell("B1", "Revenue").
//		Column("A2:A4", []any{"Widget", "Gadget", "Gizmo"}).
//		Column("B2:B4", []any{1000, 2500, 750}).
//		Cell("B5", xlreport.Formula("SUM(B2:B4)"))
//
//	err := r.SaveAs("report.xlsx")
package xlreport

import (
	"bytes"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// Report represents an Excel workbook that can be populated with data, formulas,
// and styles using a fluent API.
type Report struct {
	file   *excelize.File
	sheets map[string]*Sheet
	err    error // sticky error for fluent API
}

// New creates a new empty Excel report with a default "Sheet1".
func New() *Report {
	return &Report{
		file:   excelize.NewFile(),
		sheets: make(map[string]*Sheet),
	}
}

// Open opens an existing Excel file as a template. The original file is not
// modified; use [Report.SaveAs] to write to a new path.
func Open(path string, opts ...excelize.Options) (*Report, error) {
	var o excelize.Options
	if len(opts) > 0 {
		o = opts[0]
	}
	f, err := excelize.OpenFile(path, o)
	if err != nil {
		return nil, fmt.Errorf("xlreport: open %q: %w", path, err)
	}
	return &Report{
		file:   f,
		sheets: make(map[string]*Sheet),
	}, nil
}

// FromReader creates a report from an [io.Reader], which is useful for reading
// templates from HTTP requests, embedded files, or any other byte source.
func FromReader(r io.Reader, opts ...excelize.Options) (*Report, error) {
	var o excelize.Options
	if len(opts) > 0 {
		o = opts[0]
	}
	f, err := excelize.OpenReader(r, o)
	if err != nil {
		return nil, fmt.Errorf("xlreport: read template: %w", err)
	}
	return &Report{
		file:   f,
		sheets: make(map[string]*Sheet),
	}, nil
}

// Sheet returns a [Sheet] handle for the named sheet. If the sheet does not
// exist, it is created automatically. The returned Sheet supports method
// chaining for a fluent API.
func (r *Report) Sheet(name string) *Sheet {
	if s, ok := r.sheets[name]; ok {
		return s
	}
	// Check if the sheet already exists in the file.
	if idx, _ := r.file.GetSheetIndex(name); idx == -1 {
		if _, err := r.file.NewSheet(name); err != nil {
			r.setErr(fmt.Errorf("xlreport: create sheet %q: %w", name, err))
		}
	}
	s := &Sheet{report: r, name: name}
	r.sheets[name] = s
	return s
}

// Sheets returns the names of all sheets in the workbook, in order.
func (r *Report) Sheets() []string {
	return r.file.GetSheetList()
}

// RenameSheet renames a sheet. Returns the Report for chaining.
func (r *Report) RenameSheet(old, new string) *Report {
	if err := r.file.SetSheetName(old, new); err != nil {
		r.setErr(fmt.Errorf("xlreport: rename sheet %q to %q: %w", old, new, err))
	}
	if s, ok := r.sheets[old]; ok {
		s.name = new
		r.sheets[new] = s
		delete(r.sheets, old)
	}
	return r
}

// DeleteSheet deletes a sheet by name. Returns the Report for chaining.
func (r *Report) DeleteSheet(name string) *Report {
	if err := r.file.DeleteSheet(name); err != nil {
		r.setErr(fmt.Errorf("xlreport: delete sheet %q: %w", name, err))
	}
	delete(r.sheets, name)
	return r
}

// SetActiveSheet sets which sheet is displayed when the workbook is opened.
func (r *Report) SetActiveSheet(name string) *Report {
	idx, err := r.file.GetSheetIndex(name)
	if err != nil {
		r.setErr(fmt.Errorf("xlreport: set active sheet %q: %w", name, err))
		return r
	}
	if idx == -1 {
		r.setErr(fmt.Errorf("xlreport: set active sheet %q: sheet not found", name))
		return r
	}
	r.file.SetActiveSheet(idx)
	return r
}

// SaveAs writes the report to the specified file path.
func (r *Report) SaveAs(path string, opts ...excelize.Options) error {
	if r.err != nil {
		return r.err
	}
	var o excelize.Options
	if len(opts) > 0 {
		o = opts[0]
	}
	if err := r.file.SaveAs(path, o); err != nil {
		return fmt.Errorf("xlreport: save %q: %w", path, err)
	}
	return nil
}

// WriteTo writes the Excel file to any [io.Writer]. This is useful for
// streaming reports directly to HTTP responses, cloud storage, etc.
func (r *Report) WriteTo(w io.Writer) (int64, error) {
	if r.err != nil {
		return 0, r.err
	}
	buf, err := r.file.WriteToBuffer()
	if err != nil {
		return 0, fmt.Errorf("xlreport: write to buffer: %w", err)
	}
	n, err := w.Write(buf.Bytes())
	return int64(n), err
}

// Bytes returns the Excel file as a byte slice.
func (r *Report) Bytes() ([]byte, error) {
	if r.err != nil {
		return nil, r.err
	}
	buf, err := r.file.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("xlreport: write to buffer: %w", err)
	}
	return buf.Bytes(), nil
}

// Buffer returns the Excel file as a [*bytes.Buffer].
func (r *Report) Buffer() (*bytes.Buffer, error) {
	if r.err != nil {
		return nil, r.err
	}
	buf, err := r.file.WriteToBuffer()
	if err != nil {
		return nil, fmt.Errorf("xlreport: write to buffer: %w", err)
	}
	return buf, nil
}

// Close closes the underlying Excel file and releases resources.
// Always call Close when done with the report.
func (r *Report) Close() error {
	if r.file == nil {
		return nil
	}
	return r.file.Close()
}

// Err returns the first error encountered during fluent method chaining.
// This allows you to check for errors after a chain of operations rather
// than checking after each call.
func (r *Report) Err() error {
	return r.err
}

func (r *Report) setErr(err error) {
	if r.err == nil {
		r.err = err
	}
}
