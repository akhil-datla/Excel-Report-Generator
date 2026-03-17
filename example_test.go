package xlreport_test

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"

	"github.com/akhil-datla/xlreport"
)

func Example() {
	r := xlreport.New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "Hello, World!").
		Cell("B1", 42).
		Cell("C1", xlreport.Formula("B1*2"))

	path := filepath.Join(os.TempDir(), "example.xlsx")
	if err := r.SaveAs(path); err != nil {
		fmt.Println("error:", err)
		return
	}
	fmt.Println("Report saved!")
	// Output: Report saved!
}

func Example_table() {
	r := xlreport.New()
	defer r.Close()

	r.Sheet("Sales").
		Row("A1:C1", []any{"Product", "Q1", "Q2"}).
		Table("A2:C4", [][]any{
			{"Widget", 1500, 1800},
			{"Gadget", 2300, 2100},
			{"Gizmo", 800, 950},
		}).
		Cell("B5", xlreport.Formula("SUM(B2:B4)")).
		Cell("C5", xlreport.Formula("SUM(C2:C4)")).
		Style("A1:C1", xlreport.Bold, xlreport.BgColor("#4472C4"), xlreport.FontColor("#FFFFFF")).
		SetColWidth("A:C", 15)

	path := filepath.Join(os.TempDir(), "sales.xlsx")
	if err := r.SaveAs(path); err != nil {
		fmt.Println("error:", err)
		return
	}
	fmt.Println("Sales report saved!")
	// Output: Sales report saved!
}

func Example_writeTo() {
	r := xlreport.New()
	defer r.Close()

	r.Sheet("Sheet1").Cell("A1", "streamed")

	var buf bytes.Buffer
	if _, err := r.WriteTo(&buf); err != nil {
		fmt.Println("error:", err)
		return
	}
	fmt.Printf("Wrote %d bytes\n", buf.Len())
	// Output is non-deterministic (file size varies), so just check it's positive.
}

func Example_template() {
	// Create a "template" first for this example.
	t := xlreport.New()
	t.Sheet("Sheet1").Cell("A1", "Template Header")
	templatePath := filepath.Join(os.TempDir(), "template_example.xlsx")
	t.SaveAs(templatePath)
	t.Close()

	// Now open it as a template and add data.
	r, err := xlreport.Open(templatePath)
	if err != nil {
		fmt.Println("error:", err)
		return
	}
	defer r.Close()

	r.Sheet("Sheet1").Cell("A2", "Data added to template")

	outputPath := filepath.Join(os.TempDir(), "from_template.xlsx")
	if err := r.SaveAs(outputPath); err != nil {
		fmt.Println("error:", err)
		return
	}
	fmt.Println("Template report saved!")
	// Output: Template report saved!
}
