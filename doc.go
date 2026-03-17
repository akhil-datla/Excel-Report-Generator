// Package xlreport provides a fluent API for generating Excel reports from Go.
//
// xlreport wraps [github.com/xuri/excelize/v2] with an ergonomic, chainable
// interface that makes it simple to create professional Excel reports with
// data, formulas, and styling — in just a few lines of code.
//
// # Creating a report
//
// Use [New] to create an empty workbook, [Open] to use an existing file as a
// template, or [FromReader] to read a template from any [io.Reader]:
//
//	r := xlreport.New()
//	defer r.Close()
//
// # Adding data
//
// The [Sheet] type provides a fluent API for populating cells:
//
//	r.Sheet("Sales").
//		Cell("A1", "Product").
//		Cell("B1", "Revenue").
//		Row("A2:C2", []any{"Widget", 1500, 1800}).
//		Column("D1:D3", []any{"Total", 3300, 0}).
//		Table("A4:C6", [][]any{...})
//
// # Formulas
//
// Use [Formula] as a cell value to insert Excel formulas:
//
//	s.Cell("B5", xlreport.Formula("SUM(B2:B4)"))
//
// # Styling
//
// Combine [StyleOption] values to format cells and ranges:
//
//	s.Style("A1:D1", xlreport.Bold, xlreport.BgColor("#4472C4"), xlreport.FontColor("#FFFFFF"))
//
// # Saving
//
// Write the finished report to a file, [io.Writer], or byte slice:
//
//	r.SaveAs("report.xlsx")         // file
//	r.WriteTo(httpResponseWriter)   // stream
//	b, _ := r.Bytes()               // bytes
//
// # Error handling
//
// xlreport uses a sticky error pattern. Errors during method chaining are
// captured and can be checked once at the end:
//
//	r.Sheet("Sheet1").Cell("A1", "hi").Row("A2:C2", data)
//	if err := r.Err(); err != nil {
//		log.Fatal(err)
//	}
package xlreport
