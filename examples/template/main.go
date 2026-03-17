// Template example demonstrating how to open an existing Excel file
// and populate it with data.
package main

import (
	"bytes"
	"fmt"
	"log"

	"github.com/akhil-datla/xlreport"
)

func main() {
	// For this example, we create a "template" in memory and then reopen it.
	// In real usage, you'd open an existing .xlsx file from disk.
	tmpl := xlreport.New()
	tmpl.Sheet("Invoice").
		Cell("A1", "Invoice #").
		Cell("A2", "Date:").
		Cell("A3", "Customer:").
		Row("A5:D5", []any{"Item", "Qty", "Price", "Total"}).
		Style("A5:D5", xlreport.Bold, xlreport.BgColor("#E2EFDA"), xlreport.Border("thin", "#000000")).
		SetColWidth("A", 20).
		SetColWidth("B:D", 12)
	tmpl.DeleteSheet("Sheet1")

	buf, err := tmpl.Buffer()
	if err != nil {
		log.Fatal(err)
	}
	tmpl.Close()

	// Now open the template and fill in data.
	r, err := xlreport.FromReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()

	r.Sheet("Invoice").
		Cell("B1", "INV-2024-001").
		Cell("B2", "2024-12-15").
		Cell("B3", "Acme Corp").
		Table("A6:D8", [][]any{
			{"Widget Pro", 10, 25.00, xlreport.Formula("B6*C6")},
			{"Gadget Plus", 5, 49.99, xlreport.Formula("B7*C7")},
			{"Gizmo Ultra", 20, 12.50, xlreport.Formula("B8*C8")},
		}).
		Cell("C9", "Subtotal:").
		Cell("D9", xlreport.Formula("SUM(D6:D8)")).
		Style("C9:D9", xlreport.Bold).
		Style("C6:D9", xlreport.NumFmt("$#,##0.00"))

	if err := r.SaveAs("invoice.xlsx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("Created invoice.xlsx")
}
