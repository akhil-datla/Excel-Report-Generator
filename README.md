# Excel Report Generator

Quickly convert your raw data into Microsoft Excel reports using Go.

## Features

- Generate Excel reports from Go code.
- Input scalar, vector, and table data into specific Excel sheet ranges.
- Use your own templates for consistent styling.
- Simple, extensible API for automation and integration.

## Getting Started

### Prerequisites

- Go 1.18+
- A template Excel file (e.g., `template.xlsx`) **(OPTIONAL)**

### Installation

Clone the repository:

```bash
git clone https://github.com/akhil-datla/Excel-Report-Generator.git
cd Excel-Report-Generator
```

Install dependencies:

```bash
go mod tidy
```

### Usage

See `main.go` for an example:

```go
import (
    "main/reportgen"
    "fmt"
)

func main() {
    opts := reportgen.Options{
        TemplateFilePath: "template.xlsx", // leave empty if you want to create a blank spreadsheet
        OutputFilePath:   "output.xlsx",
        NewFile:          true, // if true, creates a new file using "OutputFilePath"
    }

    reportGenerator, err := reportgen.NewReportGenerator(opts)
    if err != nil {
        panic(err)
    }

    // Scalar data
    scalar := reportgen.Scalar{Value: "hello"}
    err = reportGenerator.InputData("Sheet1", "A1", scalar)

    // Vector data
    vector := reportgen.Vector{Values: []interface{}{1, 2, 3, 4, 5}}
    err = reportGenerator.InputData("Sheet1", "B1:B5", vector)

    // Table data
    table := reportgen.Table{
        Values: [][]interface{}{
            {1, 2, 3},
            {4, 5, 6},
            {7, 8, 9},
        },
    }
    err = reportGenerator.InputData("Sheet1", "C1:E3", table)

    // Save the sheet
    err = reportGenerator.SaveSheet()
    fmt.Println("Report generated successfully")
}
```

### API Overview

The core logic is in `reportgen/`:

- `reportgen.go`: Main report generator implementation.
- `types.go`: Defines types for Scalars, Vectors, Tables, and Options.
- `helpers.go`: Utility functions.
- `errors.go`: Custom error types.

#### Creating a Report Generator

```go
opts := reportgen.Options{ /* ... */ }
rg, err := reportgen.NewReportGenerator(opts)
```

#### Inputting Data

- **Scalar:** Single cell value
- **Vector:** 1D range (row or column)
- **Table:** 2D array for a block of cells

```go
err = rg.InputData("Sheet1", "A1", reportgen.Scalar{Value: 42})
err = rg.InputData("Sheet1", "B1:B5", reportgen.Vector{Values: []interface{}{1,2,3}})
err = rg.InputData("Sheet1", "C1:E3", reportgen.Table{Values: [][]interface{}{{1, 2, 3},{4, 5, 6},{7, 8, 9},},})
```

#### Saving the Report

```go
err = rg.SaveSheet()
```

## License

This project is licensed under the terms of the [MIT License](LICENSE).
