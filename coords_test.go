package xlreport

import "testing"

func TestParseRange(t *testing.T) {
	tests := []struct {
		input   string
		want    rangeCoords
		wantErr bool
	}{
		{"A1:C3", rangeCoords{1, 1, 3, 3}, false},
		{"B2:D5", rangeCoords{2, 2, 4, 5}, false},
		{"$A$1:$C$3", rangeCoords{1, 1, 3, 3}, false},
		// Reversed range should be normalized.
		{"C3:A1", rangeCoords{1, 1, 3, 3}, false},
		{"D1:B1", rangeCoords{2, 1, 4, 1}, false},
		// Invalid inputs.
		{"A1", rangeCoords{}, true},
		{"", rangeCoords{}, true},
		{"A1:B2:C3", rangeCoords{}, true},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := parseRange(tt.input)
			if (err != nil) != tt.wantErr {
				t.Fatalf("parseRange(%q) error = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
			if !tt.wantErr && got != tt.want {
				t.Errorf("parseRange(%q) = %+v, want %+v", tt.input, got, tt.want)
			}
		})
	}
}

func TestRangeCoordsDimensions(t *testing.T) {
	rc := rangeCoords{col1: 1, row1: 1, col2: 3, row2: 5}
	if rc.rows() != 5 {
		t.Errorf("rows() = %d, want 5", rc.rows())
	}
	if rc.cols() != 3 {
		t.Errorf("cols() = %d, want 3", rc.cols())
	}
}
