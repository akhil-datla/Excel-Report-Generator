// Basic example demonstrating core xlreport features.
package main

import (
	"fmt"
	"log"

	"github.com/akhil-datla/xlreport"
)

func main() {
	r := xlreport.New()
	defer r.Close()

	r.Sheet("Sheet1").
		Cell("A1", "Hello from xlreport!").
		Cell("A2", 42).
		Cell("A3", 3.14).
		Cell("A4", true).
		Cell("A5", xlreport.Formula("A2*A3"))

	if err := r.SaveAs("basic.xlsx"); err != nil {
		log.Fatal(err)
	}
	fmt.Println("Created basic.xlsx")
}
