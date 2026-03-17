package xlreport

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Sheet provides a fluent API for populating a single worksheet with data,
// formulas, and styles. All methods return the Sheet for chaining.
type Sheet struct {
	report *Report
	name   string
}

// Cell sets a single cell's value. The value can be a string, number, bool,
// [Formula], or any type supported by excelize.
//
//	s.Cell("A1", "Hello").Cell("B1", 42).Cell("C1", xlreport.Formula("SUM(A1:B1)"))
func (s *Sheet) Cell(cell string, value any) *Sheet {
	if s.report.err != nil {
		return s
	}
	switch v := value.(type) {
	case Formula:
		if err := s.report.file.SetCellFormula(s.name, cell, string(v)); err != nil {
			s.report.setErr(fmt.Errorf("xlreport: set formula at %s!%s: %w", s.name, cell, err))
		}
	default:
		if err := s.report.file.SetCellValue(s.name, cell, v); err != nil {
			s.report.setErr(fmt.Errorf("xlreport: set cell %s!%s: %w", s.name, cell, err))
		}
	}
	return s
}

// Column writes a slice of values vertically into a column range.
// The range must specify start and end cells (e.g., "A1:A5").
// The length of values must match the range height.
//
//	s.Column("A1:A5", []any{1, 2, 3, 4, 5})
func (s *Sheet) Column(cellRange string, values []any) *Sheet {
	if s.report.err != nil {
		return s
	}
	coords, err := parseRange(cellRange)
	if err != nil {
		s.report.setErr(err)
		return s
	}
	height := coords.rows()
	if len(values) != height {
		s.report.setErr(fmt.Errorf("xlreport: column %q expects %d values, got %d", cellRange, height, len(values)))
		return s
	}
	for i, val := range values {
		cellName, err := excelize.CoordinatesToCellName(coords.col1, coords.row1+i)
		if err != nil {
			s.report.setErr(fmt.Errorf("xlreport: column %q: %w", cellRange, err))
			return s
		}
		s.Cell(cellName, val)
	}
	return s
}

// Row writes a slice of values horizontally into a row range.
// The range must specify start and end cells (e.g., "A1:E1").
// The length of values must match the range width.
//
//	s.Row("A1:E1", []any{"Name", "Age", "City", "State", "Zip"})
func (s *Sheet) Row(cellRange string, values []any) *Sheet {
	if s.report.err != nil {
		return s
	}
	coords, err := parseRange(cellRange)
	if err != nil {
		s.report.setErr(err)
		return s
	}
	width := coords.cols()
	if len(values) != width {
		s.report.setErr(fmt.Errorf("xlreport: row %q expects %d values, got %d", cellRange, width, len(values)))
		return s
	}
	for i, val := range values {
		cellName, err := excelize.CoordinatesToCellName(coords.col1+i, coords.row1)
		if err != nil {
			s.report.setErr(fmt.Errorf("xlreport: row %q: %w", cellRange, err))
			return s
		}
		s.Cell(cellName, val)
	}
	return s
}

// Table writes a 2D slice of values into a rectangular range.
// The outer slice represents rows, inner slices represent columns.
// The dimensions must match the range exactly.
//
//	s.Table("A1:C3", [][]any{
//		{"Name", "Age", "City"},
//		{"Alice", 30, "NYC"},
//		{"Bob", 25, "LA"},
//	})
func (s *Sheet) Table(cellRange string, values [][]any) *Sheet {
	if s.report.err != nil {
		return s
	}
	coords, err := parseRange(cellRange)
	if err != nil {
		s.report.setErr(err)
		return s
	}
	height := coords.rows()
	width := coords.cols()
	if len(values) != height {
		s.report.setErr(fmt.Errorf("xlreport: table %q expects %d rows, got %d", cellRange, height, len(values)))
		return s
	}
	for i, row := range values {
		if len(row) != width {
			s.report.setErr(fmt.Errorf("xlreport: table %q row %d expects %d columns, got %d", cellRange, i, width, len(row)))
			return s
		}
		for j, val := range row {
			cellName, err := excelize.CoordinatesToCellName(coords.col1+j, coords.row1+i)
			if err != nil {
				s.report.setErr(fmt.Errorf("xlreport: table %q: %w", cellRange, err))
				return s
			}
			s.Cell(cellName, val)
		}
	}
	return s
}

// MergeCells merges a range of cells (e.g., "A1:D1").
func (s *Sheet) MergeCells(cellRange string) *Sheet {
	if s.report.err != nil {
		return s
	}
	parts := strings.Split(cellRange, ":")
	if len(parts) != 2 {
		s.report.setErr(fmt.Errorf("xlreport: merge %q: expected range like A1:D1", cellRange))
		return s
	}
	if err := s.report.file.MergeCell(s.name, parts[0], parts[1]); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: merge %s!%s: %w", s.name, cellRange, err))
	}
	return s
}

// SetColWidth sets the width of columns in the given range (e.g., "A", "B", or "A:D").
func (s *Sheet) SetColWidth(cols string, width float64) *Sheet {
	if s.report.err != nil {
		return s
	}
	parts := strings.Split(cols, ":")
	start := parts[0]
	end := start
	if len(parts) == 2 {
		end = parts[1]
	}
	if err := s.report.file.SetColWidth(s.name, start, end, width); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: set col width %s!%s: %w", s.name, cols, err))
	}
	return s
}

// SetRowHeight sets the height of a row.
func (s *Sheet) SetRowHeight(row int, height float64) *Sheet {
	if s.report.err != nil {
		return s
	}
	if err := s.report.file.SetRowHeight(s.name, row, height); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: set row height %s!%d: %w", s.name, row, err))
	}
	return s
}

// Style applies one or more [StyleOption] to a cell or range (e.g., "A1" or "A1:D1").
//
//	s.Style("A1:D1", xlreport.Bold, xlreport.FontSize(14), xlreport.BgColor("#4472C4"))
func (s *Sheet) Style(cellRange string, opts ...StyleOption) *Sheet {
	if s.report.err != nil {
		return s
	}
	if len(opts) == 0 {
		return s
	}

	cfg := &styleConfig{}
	for _, opt := range opts {
		opt(cfg)
	}

	style := cfg.toExcelizeStyle()
	styleID, err := s.report.file.NewStyle(style)
	if err != nil {
		s.report.setErr(fmt.Errorf("xlreport: create style: %w", err))
		return s
	}

	// Determine if this is a single cell or a range.
	if strings.Contains(cellRange, ":") {
		parts := strings.Split(cellRange, ":")
		if err := s.report.file.SetCellStyle(s.name, parts[0], parts[1], styleID); err != nil {
			s.report.setErr(fmt.Errorf("xlreport: apply style %s!%s: %w", s.name, cellRange, err))
		}
	} else {
		if err := s.report.file.SetCellStyle(s.name, cellRange, cellRange, styleID); err != nil {
			s.report.setErr(fmt.Errorf("xlreport: apply style %s!%s: %w", s.name, cellRange, err))
		}
	}
	return s
}

// FreezePane freezes rows and columns so they remain visible while scrolling.
// The cell parameter specifies the top-left cell of the scrollable region.
// For example, "A2" freezes the first row, "B1" freezes the first column,
// and "B2" freezes both the first row and first column.
//
//	s.FreezePane("A2")   // freeze top row
//	s.FreezePane("B1")   // freeze first column
//	s.FreezePane("B3")   // freeze first column and top 2 rows
func (s *Sheet) FreezePane(cell string) *Sheet {
	if s.report.err != nil {
		return s
	}
	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		s.report.setErr(fmt.Errorf("xlreport: freeze pane %q: %w", cell, err))
		return s
	}
	panes := excelize.Panes{
		Freeze:      true,
		TopLeftCell: cell,
		XSplit:      col - 1,
		YSplit:      row - 1,
	}
	// Set the active pane based on what's frozen.
	switch {
	case col > 1 && row > 1:
		panes.ActivePane = "bottomRight"
	case row > 1:
		panes.ActivePane = "bottomLeft"
	case col > 1:
		panes.ActivePane = "topRight"
	}
	if err := s.report.file.SetPanes(s.name, &panes); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: freeze pane %s!%s: %w", s.name, cell, err))
	}
	return s
}

// AutoFilter adds an auto-filter (dropdown arrows) to a range. This is
// typically applied to a header row so users can filter data in Excel.
//
//	s.AutoFilter("A1:E1")   // filter arrows on columns A through E, row 1
func (s *Sheet) AutoFilter(cellRange string) *Sheet {
	if s.report.err != nil {
		return s
	}
	if err := s.report.file.AutoFilter(s.name, cellRange, nil); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: auto filter %s!%s: %w", s.name, cellRange, err))
	}
	return s
}

// AddImage inserts an image from a file path at the specified cell.
// Supported formats: PNG, JPG, GIF, TIFF, BMP, SVG, EMF, WMF.
//
//	s.AddImage("A1", "logo.png")
func (s *Sheet) AddImage(cell, path string) *Sheet {
	if s.report.err != nil {
		return s
	}
	if err := s.report.file.AddPicture(s.name, cell, path, nil); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: add image %s!%s %q: %w", s.name, cell, path, err))
	}
	return s
}

// AddImageBytes inserts an image from raw bytes at the specified cell.
// The extension should include the dot (e.g., ".png", ".jpg").
//
//	s.AddImageBytes("A1", ".png", logoBytes)
func (s *Sheet) AddImageBytes(cell, extension string, data []byte) *Sheet {
	if s.report.err != nil {
		return s
	}
	pic := &excelize.Picture{
		Extension: extension,
		File:      data,
	}
	if err := s.report.file.AddPictureFromBytes(s.name, cell, pic); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: add image bytes %s!%s: %w", s.name, cell, err))
	}
	return s
}

// SetVisible controls whether the sheet is visible or hidden in Excel.
//
//	s.SetVisible(false) // hide the sheet
func (s *Sheet) SetVisible(visible bool) *Sheet {
	if s.report.err != nil {
		return s
	}
	if err := s.report.file.SetSheetVisible(s.name, visible); err != nil {
		s.report.setErr(fmt.Errorf("xlreport: set sheet visible %s: %w", s.name, err))
	}
	return s
}

// Name returns the name of this sheet.
func (s *Sheet) Name() string {
	return s.name
}
