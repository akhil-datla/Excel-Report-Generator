// Styled example demonstrating formatting, tables, and formulas.
package main

import (
	"fmt"
	"log"

	"github.com/akhil-datla/xlreport"
)

func main() {
	r := xlreport.New()
	defer r.Close()

	// Delete default sheet, create our own.
	r.Sheet("Sales Report")
	r.DeleteSheet("Sheet1")

	s := r.Sheet("Sales Report")

	// Title
	s.Cell("A1", "Quarterly Sales Report").
		MergeCells("A1:E1").
		Style("A1:E1", xlreport.Bold, xlreport.FontSize(18), xlreport.AlignCenter,
			xlreport.BgColor("#1F4E79"), xlreport.FontColor("#FFFFFF"))

	// Headers
	s.Row("A3:E3", []any{"Product", "Q1", "Q2", "Q3", "Q4"}).
		Style("A3:E3", xlreport.Bold, xlreport.BgColor("#4472C4"), xlreport.FontColor("#FFFFFF"),
			xlreport.AlignCenter, xlreport.Border("thin", "#000000"))

	// Data
	s.Table("A4:E7", [][]any{
		{"Widget Pro", 15200, 18400, 21000, 24500},
		{"Gadget Plus", 23000, 21500, 25800, 28000},
		{"Gizmo Ultra", 8500, 9200, 11000, 13500},
		{"Thingamajig", 4200, 5100, 4800, 6200},
	})

	// Totals row with formulas
	s.Cell("A8", "TOTAL").
		Cell("B8", xlreport.Formula("SUM(B4:B7)")).
		Cell("C8", xlreport.Formula("SUM(C4:C7)")).
		Cell("D8", xlreport.Formula("SUM(D4:D7)")).
		Cell("E8", xlreport.Formula("SUM(E4:E7)")).
		Style("A8:E8", xlreport.Bold, xlreport.BgColor("#D6E4F0"),
			xlreport.Border("medium", "#000000"))

	// Format numbers as currency
	s.Style("B4:E8", xlreport.NumFmt("#,##0"))

	// Column widths
	s.SetColWidth("A", 18).
		SetColWidth("B:E", 14).
		SetRowHeight(1, 30)

	if err := r.SaveAs("sales_report.xlsx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("Created sales_report.xlsx")
}
