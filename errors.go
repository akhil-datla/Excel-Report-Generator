package xlreport

import "errors"

// Sentinel errors returned by xlreport functions.
var (
	// ErrClosed is returned when operating on a closed report.
	ErrClosed = errors.New("xlreport: report is closed")
)
