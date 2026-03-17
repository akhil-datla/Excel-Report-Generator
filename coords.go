package xlreport

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// rangeCoords represents normalized cell range coordinates.
type rangeCoords struct {
	col1, row1 int // top-left corner
	col2, row2 int // bottom-right corner
}

func (rc rangeCoords) rows() int { return rc.row2 - rc.row1 + 1 }
func (rc rangeCoords) cols() int { return rc.col2 - rc.col1 + 1 }

// parseRange parses an Excel range reference (e.g., "A1:C5") into normalized
// coordinates. Handles reversed ranges like "C5:A1" by swapping.
func parseRange(ref string) (rangeCoords, error) {
	ref = strings.ReplaceAll(ref, "$", "")
	parts := strings.Split(ref, ":")
	if len(parts) != 2 {
		return rangeCoords{}, fmt.Errorf("xlreport: invalid range %q: expected format like A1:C5", ref)
	}

	col1, row1, err := excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		return rangeCoords{}, fmt.Errorf("xlreport: invalid cell %q in range %q: %w", parts[0], ref, err)
	}
	col2, row2, err := excelize.CellNameToCoordinates(parts[1])
	if err != nil {
		return rangeCoords{}, fmt.Errorf("xlreport: invalid cell %q in range %q: %w", parts[1], ref, err)
	}

	// Normalize: ensure top-left <= bottom-right.
	if col2 < col1 {
		col1, col2 = col2, col1
	}
	if row2 < row1 {
		row1, row2 = row2, row1
	}

	return rangeCoords{col1: col1, row1: row1, col2: col2, row2: row2}, nil
}
