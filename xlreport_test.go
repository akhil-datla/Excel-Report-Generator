package xlreport

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

func TestNew(t *testing.T) {
	r := New()
	defer r.Close()

	sheets := r.Sheets()
	if len(sheets) != 1 || sheets[0] != "Sheet1" {
		t.Errorf("New() should create report with Sheet1, got %v", sheets)
	}
}

func TestNewAndSaveAs(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "hello")

	path := filepath.Join(t.TempDir(), "test.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	if _, err := os.Stat(path); err != nil {
		t.Fatalf("output file not created: %v", err)
	}
}

func TestOpenAndRoundTrip(t *testing.T) {
	// Create a file first.
	r := New()
	r.Sheet("Sheet1").Cell("A1", "test value")
	path := filepath.Join(t.TempDir(), "template.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	r.Close()

	// Open it as a template.
	r2, err := Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer r2.Close()

	// Verify we can read the value back.
	val, err := r2.file.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if val != "test value" {
		t.Errorf("expected %q, got %q", "test value", val)
	}
}

func TestFromReader(t *testing.T) {
	// Create a file in memory.
	r := New()
	r.Sheet("Sheet1").Cell("A1", 42)
	buf, err := r.Buffer()
	if err != nil {
		t.Fatalf("Buffer: %v", err)
	}
	r.Close()

	// Open from reader.
	r2, err := FromReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("FromReader: %v", err)
	}
	defer r2.Close()

	val, err := r2.file.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if val != "42" {
		t.Errorf("expected %q, got %q", "42", val)
	}
}

func TestWriteTo(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "hello")

	var buf bytes.Buffer
	n, err := r.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	if n == 0 {
		t.Error("WriteTo wrote 0 bytes")
	}
	if buf.Len() == 0 {
		t.Error("buffer is empty")
	}
}

func TestBytes(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "hello")

	b, err := r.Bytes()
	if err != nil {
		t.Fatalf("Bytes: %v", err)
	}
	if len(b) == 0 {
		t.Error("Bytes returned empty slice")
	}
}

func TestCellScalar(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "text").
		Cell("B1", 42).
		Cell("C1", 3.14).
		Cell("D1", true)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	tests := []struct {
		cell string
		want string
	}{
		{"A1", "text"},
		{"B1", "42"},
		{"C1", "3.14"},
		{"D1", "TRUE"},
	}
	for _, tt := range tests {
		val, err := r.file.GetCellValue("Sheet1", tt.cell)
		if err != nil {
			t.Errorf("GetCellValue(%s): %v", tt.cell, err)
			continue
		}
		if val != tt.want {
			t.Errorf("cell %s: got %q, want %q", tt.cell, val, tt.want)
		}
	}
}

func TestCellFormula(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", 10).
		Cell("A2", 20).
		Cell("A3", Formula("SUM(A1:A2)"))

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula, err := r.file.GetCellFormula("Sheet1", "A3")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if formula != "SUM(A1:A2)" {
		t.Errorf("formula: got %q, want %q", formula, "SUM(A1:A2)")
	}
}

func TestColumn(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("A1:A5", []any{10, 20, 30, 40, 50})

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	for i, want := range []string{"10", "20", "30", "40", "50"} {
		cell := "A" + string(rune('1'+i))
		val, err := r.file.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Errorf("GetCellValue(%s): %v", cell, err)
			continue
		}
		if val != want {
			t.Errorf("cell %s: got %q, want %q", cell, val, want)
		}
	}
}

func TestColumnSizeMismatch(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("A1:A3", []any{1, 2, 3, 4})

	if r.Err() == nil {
		t.Error("expected error for size mismatch, got nil")
	}
}

func TestRow(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Row("A1:E1", []any{"a", "b", "c", "d", "e"})

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	for i, want := range []string{"a", "b", "c", "d", "e"} {
		col := string(rune('A' + i))
		cell := col + "1"
		val, err := r.file.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Errorf("GetCellValue(%s): %v", cell, err)
			continue
		}
		if val != want {
			t.Errorf("cell %s: got %q, want %q", cell, val, want)
		}
	}
}

func TestRowSizeMismatch(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Row("A1:C1", []any{1, 2})

	if r.Err() == nil {
		t.Error("expected error for size mismatch, got nil")
	}
}

func TestTable(t *testing.T) {
	r := New()
	defer r.Close()

	data := [][]any{
		{1, 2, 3},
		{4, 5, 6},
	}
	r.Sheet("Sheet1").Table("A1:C2", data)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	expected := map[string]string{
		"A1": "1", "B1": "2", "C1": "3",
		"A2": "4", "B2": "5", "C2": "6",
	}
	for cell, want := range expected {
		val, err := r.file.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Errorf("GetCellValue(%s): %v", cell, err)
			continue
		}
		if val != want {
			t.Errorf("cell %s: got %q, want %q", cell, val, want)
		}
	}
}

func TestTableSizeMismatch(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Table("A1:C2", [][]any{
		{1, 2, 3},
	})

	if r.Err() == nil {
		t.Error("expected error for row count mismatch, got nil")
	}
}

func TestTableColumnMismatch(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Table("A1:C2", [][]any{
		{1, 2},
		{3, 4},
	})

	if r.Err() == nil {
		t.Error("expected error for column count mismatch, got nil")
	}
}

func TestMergeCells(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "Merged Header").
		MergeCells("A1:D1")

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Verify file can be saved (merge is valid).
	path := filepath.Join(t.TempDir(), "merge.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestSheetAutoCreate(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("NewSheet").Cell("A1", "test")

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	found := false
	for _, name := range r.Sheets() {
		if name == "NewSheet" {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("sheet 'NewSheet' not found in %v", r.Sheets())
	}
}

func TestRenameSheet(t *testing.T) {
	r := New()
	defer r.Close()

	r.RenameSheet("Sheet1", "Data")

	sheets := r.Sheets()
	if len(sheets) != 1 || sheets[0] != "Data" {
		t.Errorf("after rename, expected [Data], got %v", sheets)
	}
}

func TestDeleteSheet(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("ToDelete")
	r.DeleteSheet("ToDelete")

	for _, name := range r.Sheets() {
		if name == "ToDelete" {
			t.Error("sheet 'ToDelete' should have been deleted")
		}
	}
}

func TestStyle(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "Styled").
		Style("A1", Bold, Italic, FontSize(14), FontColor("#FF0000"), BgColor("#FFFF00")).
		Style("A2:A5", AlignCenter, WrapText, Border("thin", "#000000"))

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Verify the file can be saved with styles.
	path := filepath.Join(t.TempDir(), "styled.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestSetColWidth(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		SetColWidth("A", 20).
		SetColWidth("B:D", 15)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestSetRowHeight(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").SetRowHeight(1, 30)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	// Trigger an error with an invalid range.
	r.Sheet("Sheet1").Column("invalid", []any{1})

	if r.Err() == nil {
		t.Fatal("expected error, got nil")
	}

	// Subsequent operations should be no-ops (sticky error).
	r.Sheet("Sheet1").Cell("A1", "should not be set")

	// Error should still be the original one.
	if r.Err() == nil {
		t.Fatal("sticky error should persist")
	}
}

func TestCloseIdempotent(t *testing.T) {
	r := New()
	if err := r.Close(); err != nil {
		t.Fatalf("first Close: %v", err)
	}
	// Second close should not panic.
	r.Close()
}

func TestOpenNonExistent(t *testing.T) {
	_, err := Open("/nonexistent/path.xlsx")
	if err == nil {
		t.Error("expected error opening nonexistent file")
	}
}

func TestFluentChaining(t *testing.T) {
	r := New()
	defer r.Close()

	// This should compile and work — verifying the fluent API chains correctly.
	r.Sheet("Sheet1").
		Cell("A1", "Name").
		Cell("B1", "Age").
		Cell("C1", "Score").
		Row("A2:C2", []any{"Alice", 30, 95.5}).
		Row("A3:C3", []any{"Bob", 25, 87.0}).
		Column("D1:D3", []any{"Total", Formula("B2+C2"), Formula("B3+C3")}).
		Table("E1:F2", [][]any{{"X", "Y"}, {1, 2}}).
		Style("A1:C1", Bold, BgColor("#4472C4"), FontColor("#FFFFFF")).
		SetColWidth("A:C", 15).
		MergeCells("E1:F1")

	if err := r.Err(); err != nil {
		t.Fatalf("fluent chain error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "fluent.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestReversedRange(t *testing.T) {
	r := New()
	defer r.Close()

	// C1:A1 should be normalized to A1:C1.
	r.Sheet("Sheet1").Row("C1:A1", []any{1, 2, 3})

	if err := r.Err(); err != nil {
		t.Fatalf("reversed range should be handled: %v", err)
	}

	val, _ := r.file.GetCellValue("Sheet1", "A1")
	if val != "1" {
		t.Errorf("A1: got %q, want %q", val, "1")
	}
}

func TestNumFmtStyle(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", 0.156).
		Style("A1", NumFmt("0.00%"))

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestSheetName(t *testing.T) {
	r := New()
	defer r.Close()

	s := r.Sheet("MySheet")
	if s.Name() != "MySheet" {
		t.Errorf("Name() = %q, want %q", s.Name(), "MySheet")
	}
}

func TestFontNameStyle(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "Arial text").
		Style("A1", FontName("Arial"), FontSize(12))

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "fontname.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestAllAlignmentStyles(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "left").Style("A1", AlignLeft).
		Cell("A2", "center").Style("A2", AlignCenter).
		Cell("A3", "right").Style("A3", AlignRight).
		Cell("A4", "top").Style("A4", VAlignTop).
		Cell("A5", "middle").Style("A5", VAlignMiddle).
		Cell("A6", "bottom").Style("A6", VAlignBottom)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestUnderlineAndStrikethrough(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "underlined").Style("A1", Underline).
		Cell("A2", "struck").Style("A2", Strikethrough)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestBorderStyles(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Style("A1", Border("thin", "#000000")).
		Style("A2", Border("medium", "#333333")).
		Style("A3", Border("thick", "#666666")).
		Style("A4", Border("unknown_style", "#999999")) // should default to thin

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestMergeCellsInvalidRange(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").MergeCells("A1") // no colon, invalid

	if r.Err() == nil {
		t.Error("expected error for invalid merge range, got nil")
	}
}

func TestStyleNoOpts(t *testing.T) {
	r := New()
	defer r.Close()

	// Style with no options should be a no-op, not an error.
	r.Sheet("Sheet1").Style("A1")

	if err := r.Err(); err != nil {
		t.Fatalf("Style with no opts should not error: %v", err)
	}
}

func TestRenameSheetUpdatesCache(t *testing.T) {
	r := New()
	defer r.Close()

	// Access sheet to cache it.
	r.Sheet("Sheet1").Cell("A1", "before rename")

	// Rename and verify the cached sheet handle still works.
	r.RenameSheet("Sheet1", "Renamed")
	r.Sheet("Renamed").Cell("A2", "after rename")

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val, _ := r.file.GetCellValue("Renamed", "A1")
	if val != "before rename" {
		t.Errorf("A1: got %q, want %q", val, "before rename")
	}
	val, _ = r.file.GetCellValue("Renamed", "A2")
	if val != "after rename" {
		t.Errorf("A2: got %q, want %q", val, "after rename")
	}
}

func TestSaveAsWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	// Trigger a sticky error.
	r.Sheet("Sheet1").Column("bad", []any{1})

	// SaveAs should return the sticky error.
	if err := r.SaveAs(filepath.Join(t.TempDir(), "out.xlsx")); err == nil {
		t.Error("SaveAs should return sticky error, got nil")
	}
}

func TestWriteToWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})

	var buf bytes.Buffer
	_, err := r.WriteTo(&buf)
	if err == nil {
		t.Error("WriteTo should return sticky error, got nil")
	}
}

func TestBytesWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})

	_, err := r.Bytes()
	if err == nil {
		t.Error("Bytes should return sticky error, got nil")
	}
}

func TestBufferWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})

	_, err := r.Buffer()
	if err == nil {
		t.Error("Buffer should return sticky error, got nil")
	}
}

func TestFromReaderInvalid(t *testing.T) {
	_, err := FromReader(bytes.NewReader([]byte("not a valid xlsx")))
	if err == nil {
		t.Error("expected error from invalid reader data")
	}
}

func TestCloseNilFile(t *testing.T) {
	r := &Report{file: nil}
	if err := r.Close(); err != nil {
		t.Errorf("Close on nil file should return nil, got %v", err)
	}
}

func TestColumnWithFormulas(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", 10).
		Cell("A2", 20).
		Column("B1:B2", []any{Formula("A1*2"), Formula("A2*2")})

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	f1, _ := r.file.GetCellFormula("Sheet1", "B1")
	f2, _ := r.file.GetCellFormula("Sheet1", "B2")
	if f1 != "A1*2" {
		t.Errorf("B1 formula: got %q, want %q", f1, "A1*2")
	}
	if f2 != "A2*2" {
		t.Errorf("B2 formula: got %q, want %q", f2, "A2*2")
	}
}

func TestRowWithFormulas(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Row("A1:C1", []any{10, 20, 30}).
		Row("A2:C2", []any{Formula("A1*2"), Formula("B1*2"), Formula("C1*2")})

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	for i, want := range []string{"A1*2", "B1*2", "C1*2"} {
		col := string(rune('A' + i))
		f, _ := r.file.GetCellFormula("Sheet1", col+"2")
		if f != want {
			t.Errorf("%s2 formula: got %q, want %q", col, f, want)
		}
	}
}

func TestTableWithFormulas(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Table("A1:B2", [][]any{
		{1, 2},
		{Formula("A1+B1"), Formula("A1*B1")},
	})

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	f1, _ := r.file.GetCellFormula("Sheet1", "A2")
	f2, _ := r.file.GetCellFormula("Sheet1", "B2")
	if f1 != "A1+B1" {
		t.Errorf("A2 formula: got %q, want %q", f1, "A1+B1")
	}
	if f2 != "A1*B1" {
		t.Errorf("B2 formula: got %q, want %q", f2, "A1*B1")
	}
}

func TestStickyErrorBlocksAllSheetOps(t *testing.T) {
	r := New()
	defer r.Close()

	// Trigger error.
	r.Sheet("Sheet1").Column("bad", []any{1})

	// All these should be no-ops due to sticky error.
	s := r.Sheet("Sheet1")
	s.Row("A1:C1", []any{1, 2, 3})
	s.Table("A1:B2", [][]any{{1, 2}, {3, 4}})
	s.MergeCells("A1:B1")
	s.SetColWidth("A", 20)
	s.SetRowHeight(1, 30)
	s.Style("A1", Bold)

	if r.Err() == nil {
		t.Error("sticky error should persist through all operations")
	}
}

func TestSheetCachingReturnsExisting(t *testing.T) {
	r := New()
	defer r.Close()

	s1 := r.Sheet("Sheet1")
	s2 := r.Sheet("Sheet1")

	if s1 != s2 {
		t.Error("Sheet() should return the same *Sheet instance for the same name")
	}
}

func TestMultipleSheets(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "first")
	r.Sheet("Data").Cell("A1", "second")
	r.Sheet("Summary").Cell("A1", "third")

	sheets := r.Sheets()
	if len(sheets) != 3 {
		t.Errorf("expected 3 sheets, got %d: %v", len(sheets), sheets)
	}

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "multi.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

// --- FreezePane tests ---

func TestFreezePaneTopRow(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Row("A1:C1", []any{"Name", "Age", "City"}).
		FreezePane("A2") // freeze top row

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "freeze_row.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestFreezePaneFirstColumn(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").FreezePane("B1") // freeze first column

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestFreezePaneBoth(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").FreezePane("B2") // freeze first row and first column

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestFreezePaneInvalidCell(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").FreezePane("!!!invalid")

	if r.Err() == nil {
		t.Error("expected error for invalid cell, got nil")
	}
}

func TestFreezePaneWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1}) // trigger error
	r.Sheet("Sheet1").FreezePane("A2")         // should be no-op

	if r.Err() == nil {
		t.Error("sticky error should persist")
	}
}

// --- AutoFilter tests ---

func TestAutoFilter(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").
		Row("A1:C1", []any{"Name", "Age", "City"}).
		Row("A2:C2", []any{"Alice", 30, "NYC"}).
		Row("A3:C3", []any{"Bob", 25, "LA"}).
		AutoFilter("A1:C3")

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "autofilter.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestAutoFilterWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})
	r.Sheet("Sheet1").AutoFilter("A1:C1")

	if r.Err() == nil {
		t.Error("sticky error should persist")
	}
}

// --- AddImage tests ---

func testPNGBytes(t *testing.T) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, 10, 10))
	for x := 0; x < 10; x++ {
		for y := 0; y < 10; y++ {
			img.Set(x, y, color.RGBA{R: 255, A: 255})
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatalf("encode png: %v", err)
	}
	return buf.Bytes()
}

func TestAddImage(t *testing.T) {
	r := New()
	defer r.Close()

	// Write a real PNG to a temp file.
	pngData := testPNGBytes(t)
	imgPath := filepath.Join(t.TempDir(), "test.png")
	if err := os.WriteFile(imgPath, pngData, 0644); err != nil {
		t.Fatalf("write image: %v", err)
	}

	r.Sheet("Sheet1").AddImage("A1", imgPath)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "with_image.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestAddImageNonExistent(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").AddImage("A1", "/nonexistent/logo.png")

	if r.Err() == nil {
		t.Error("expected error for nonexistent image, got nil")
	}
}

func TestAddImageWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})
	r.Sheet("Sheet1").AddImage("A1", "logo.png")

	if r.Err() == nil {
		t.Error("sticky error should persist")
	}
}

func TestAddImageBytes(t *testing.T) {
	r := New()
	defer r.Close()

	pngData := testPNGBytes(t)
	r.Sheet("Sheet1").AddImageBytes("A1", ".png", pngData)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "with_image_bytes.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestAddImageBytesWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})
	r.Sheet("Sheet1").AddImageBytes("A1", ".png", []byte{})

	if r.Err() == nil {
		t.Error("sticky error should persist")
	}
}

// --- SetVisible tests ---

func TestSetVisible(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Hidden").SetVisible(false)
	r.Sheet("Visible").SetVisible(true)

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "visibility.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestSetVisibleWithStickyError(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Column("bad", []any{1})
	r.Sheet("Sheet1").SetVisible(false)

	if r.Err() == nil {
		t.Error("sticky error should persist")
	}
}

// --- SetActiveSheet tests ---

func TestSetActiveSheet(t *testing.T) {
	r := New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "first")
	r.Sheet("Data").Cell("A1", "second")
	r.SetActiveSheet("Data")

	if err := r.Err(); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	path := filepath.Join(t.TempDir(), "active.xlsx")
	if err := r.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
}

func TestSetActiveSheetNotFound(t *testing.T) {
	r := New()
	defer r.Close()

	r.SetActiveSheet("NonExistent")

	if r.Err() == nil {
		t.Error("expected error for nonexistent sheet, got nil")
	}
}

func TestSetActiveSheetChaining(t *testing.T) {
	r := New()
	defer r.Close()

	// Verify SetActiveSheet returns *Report for chaining.
	r.Sheet("Data").Cell("A1", "test")
	r.SetActiveSheet("Data").DeleteSheet("Sheet1")

	sheets := r.Sheets()
	if len(sheets) != 1 || sheets[0] != "Data" {
		t.Errorf("expected [Data], got %v", sheets)
	}
}
