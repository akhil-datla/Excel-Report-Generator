# xlreport

[![Go Reference](https://pkg.go.dev/badge/github.com/akhil-datla/xlreport.svg)](https://pkg.go.dev/github.com/akhil-datla/xlreport)
[![Go Report Card](https://goreportcard.com/badge/github.com/akhil-datla/xlreport)](https://goreportcard.com/report/github.com/akhil-datla/xlreport)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A fluent Go library for generating Excel reports. Write data, formulas, and styles with a clean, chainable API.

```go
r := xlreport.New()
defer r.Close()

r.Sheet("Sales").
    Row("A1:C1", []any{"Product", "Q1", "Q2"}).
    Table("A2:C4", [][]any{
        {"Widget", 1500, 1800},
        {"Gadget", 2300, 2100},
        {"Gizmo",  800,  950},
    }).
    Cell("B5", xlreport.Formula("SUM(B2:B4)")).
    Style("A1:C1", xlreport.Bold, xlreport.BgColor("#4472C4"), xlreport.FontColor("#FFFFFF"))

r.SaveAs("report.xlsx")
```

## Features

- **Fluent API** — chain `Cell`, `Row`, `Column`, `Table`, `Style`, and more
- **Formulas** — set Excel formulas with `xlreport.Formula("SUM(A1:A10)")`
- **Styling** — bold, italic, colors, borders, alignment, number formats, and more
- **Freeze panes** — keep headers visible while scrolling
- **Auto-filter** — add dropdown filter arrows to data tables
- **Images** — embed PNG/JPG/GIF images from files or byte slices
- **Templates** — open existing `.xlsx` files and populate them with data
- **Streaming** — write directly to `io.Writer` for HTTP responses, S3 uploads, etc.
- **Sheet management** — create, rename, delete, hide, and reorder sheets
- **Error handling** — sticky errors let you check once after a chain of operations
- **Zero config** — `xlreport.New()` gives you a ready-to-use report

## Installation

```bash
go get github.com/akhil-datla/xlreport
```

Requires Go 1.21+.

## Quick Start

### Create a report from scratch

```go
package main

import (
    "log"
    "github.com/akhil-datla/xlreport"
)

func main() {
    r := xlreport.New()
    defer r.Close()

    r.Sheet("Sheet1").
        Cell("A1", "Name").
        Cell("B1", "Score").
        Row("A2:B2", []any{"Alice", 95}).
        Row("A3:B3", []any{"Bob", 87}).
        Cell("B4", xlreport.Formula("AVERAGE(B2:B3)")).
        Style("A1:B1", xlreport.Bold)

    if err := r.SaveAs("scores.xlsx"); err != nil {
        log.Fatal(err)
    }
}
```

### Use a template

```go
r, err := xlreport.Open("template.xlsx")
if err != nil {
    log.Fatal(err)
}
defer r.Close()

r.Sheet("Data").
    Cell("B2", "Filled from Go!").
    Table("A5:C7", myData)

r.SaveAs("filled.xlsx")
```

### Stream to HTTP response

```go
func handler(w http.ResponseWriter, req *http.Request) {
    r := xlreport.New()
    defer r.Close()

    r.Sheet("Export").
        Row("A1:C1", []any{"ID", "Name", "Email"}).
        Table("A2:C4", userData)

    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition", `attachment; filename="export.xlsx"`)
    r.WriteTo(w)
}
```

### Read template from embedded files

```go
//go:embed templates/report.xlsx
var templateFS embed.FS

func generateReport() (*bytes.Buffer, error) {
    f, _ := templateFS.Open("templates/report.xlsx")
    defer f.Close()

    r, err := xlreport.FromReader(f)
    if err != nil {
        return nil, err
    }
    defer r.Close()

    r.Sheet("Data").Table("A2:E100", records)
    return r.Buffer()
}
```

## API Reference

### Report

| Method | Description |
|--------|-------------|
| `New()` | Create a new empty report |
| `Open(path)` | Open an existing `.xlsx` file as a template |
| `FromReader(r)` | Create a report from an `io.Reader` |
| `Sheet(name)` | Get or create a sheet (returns `*Sheet` for chaining) |
| `Sheets()` | List all sheet names |
| `SetActiveSheet(name)` | Set which sheet is displayed when opened |
| `RenameSheet(old, new)` | Rename a sheet |
| `DeleteSheet(name)` | Delete a sheet |
| `SaveAs(path)` | Save to a file |
| `WriteTo(w)` | Write to an `io.Writer` |
| `Bytes()` | Get the file as `[]byte` |
| `Buffer()` | Get the file as `*bytes.Buffer` |
| `Err()` | Get the first error from chained operations |
| `Close()` | Release resources |

### Sheet

| Method | Description |
|--------|-------------|
| `Cell(cell, value)` | Set a single cell (`"A1"`, value or `Formula`) |
| `Row(range, values)` | Write values horizontally (`"A1:E1"`) |
| `Column(range, values)` | Write values vertically (`"A1:A5"`) |
| `Table(range, values)` | Write a 2D grid (`"A1:C3"`) |
| `MergeCells(range)` | Merge a range of cells (`"A1:D1"`) |
| `FreezePane(cell)` | Freeze rows/columns (`"A2"` freezes top row) |
| `AutoFilter(range)` | Add dropdown filter arrows (`"A1:E1"`) |
| `AddImage(cell, path)` | Insert image from file path |
| `AddImageBytes(cell, ext, data)` | Insert image from byte slice |
| `SetColWidth(cols, w)` | Set column width (`"A"` or `"A:D"`) |
| `SetRowHeight(row, h)` | Set row height |
| `SetVisible(bool)` | Show or hide the sheet |
| `Style(range, opts...)` | Apply styles to a cell or range |
| `Name()` | Get the sheet name |

### Styles

```go
// Font
xlreport.Bold
xlreport.Italic
xlreport.Underline
xlreport.Strikethrough
xlreport.FontName("Arial")
xlreport.FontSize(14)
xlreport.FontColor("#FF0000")

// Fill
xlreport.BgColor("#4472C4")

// Alignment
xlreport.AlignLeft
xlreport.AlignCenter
xlreport.AlignRight
xlreport.VAlignTop
xlreport.VAlignMiddle
xlreport.VAlignBottom
xlreport.WrapText

// Format
xlreport.NumFmt("#,##0.00")
xlreport.NumFmt("$#,##0.00")
xlreport.NumFmt("0%")
xlreport.NumFmt("yyyy-mm-dd")

// Borders
xlreport.Border("thin", "#000000")
xlreport.Border("medium", "#333333")
xlreport.Border("thick", "#000000")
```

### Formulas

```go
xlreport.Formula("SUM(A1:A10)")
xlreport.Formula("AVERAGE(B2:B100)")
xlreport.Formula("IF(A1>0,\"Yes\",\"No\")")
xlreport.Formula("VLOOKUP(A1,Data!A:B,2,FALSE)")
```

## Error Handling

xlreport uses a **sticky error** pattern. Errors are captured during method chaining and can be checked after:

```go
r.Sheet("Sheet1").
    Cell("A1", "hello").
    Column("B1:B5", data).
    Style("A1", xlreport.Bold)

if err := r.Err(); err != nil {
    log.Fatal(err)  // first error in the chain
}
```

All errors are wrapped with context (e.g., `xlreport: column "B1:B5" expects 5 values, got 3`) to make debugging straightforward.

## Examples

See the [examples/](examples/) directory for complete, runnable programs:

- [**basic**](examples/basic/) — simple cell, column, and formula usage
- [**styled**](examples/styled/) — formatted sales report with headers, formulas, and styling
- [**template**](examples/template/) — invoice generation from a template

## License

[MIT](LICENSE)
