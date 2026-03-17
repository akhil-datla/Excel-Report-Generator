package xlreport

import "github.com/xuri/excelize/v2"

// StyleOption configures a cell or range style. Combine multiple options
// when calling [Sheet.Style]:
//
//	s.Style("A1:D1", xlreport.Bold, xlreport.FontSize(14), xlreport.BgColor("#4472C4"))
type StyleOption func(*styleConfig)

type styleConfig struct {
	bold      bool
	italic    bool
	underline bool
	strike    bool
	fontName  string
	fontSize  float64
	fontColor string
	bgColor   string
	hAlign    string
	vAlign    string
	numFmt    string
	wrapText  bool
	border    []excelize.Border
}

func (c *styleConfig) toExcelizeStyle() *excelize.Style {
	s := &excelize.Style{}

	// Font
	font := &excelize.Font{}
	hasFont := false
	if c.bold {
		font.Bold = true
		hasFont = true
	}
	if c.italic {
		font.Italic = true
		hasFont = true
	}
	if c.underline {
		font.Underline = "single"
		hasFont = true
	}
	if c.strike {
		font.Strike = true
		hasFont = true
	}
	if c.fontName != "" {
		font.Family = c.fontName
		hasFont = true
	}
	if c.fontSize > 0 {
		font.Size = c.fontSize
		hasFont = true
	}
	if c.fontColor != "" {
		font.Color = c.fontColor
		hasFont = true
	}
	if hasFont {
		s.Font = font
	}

	// Fill (background color)
	if c.bgColor != "" {
		s.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{c.bgColor},
			Pattern: 1,
		}
	}

	// Alignment
	align := &excelize.Alignment{}
	hasAlign := false
	if c.hAlign != "" {
		align.Horizontal = c.hAlign
		hasAlign = true
	}
	if c.vAlign != "" {
		align.Vertical = c.vAlign
		hasAlign = true
	}
	if c.wrapText {
		align.WrapText = true
		hasAlign = true
	}
	if hasAlign {
		s.Alignment = align
	}

	// Number format
	if c.numFmt != "" {
		s.CustomNumFmt = &c.numFmt
	}

	// Borders
	if len(c.border) > 0 {
		s.Border = c.border
	}

	return s
}

// Bold makes text bold.
var Bold StyleOption = func(c *styleConfig) { c.bold = true }

// Italic makes text italic.
var Italic StyleOption = func(c *styleConfig) { c.italic = true }

// Underline adds a single underline to text.
var Underline StyleOption = func(c *styleConfig) { c.underline = true }

// Strikethrough adds a strikethrough to text.
var Strikethrough StyleOption = func(c *styleConfig) { c.strike = true }

// WrapText enables text wrapping in cells.
var WrapText StyleOption = func(c *styleConfig) { c.wrapText = true }

// FontName sets the font family (e.g., "Arial", "Calibri").
func FontName(name string) StyleOption {
	return func(c *styleConfig) { c.fontName = name }
}

// FontSize sets the font size in points.
func FontSize(size float64) StyleOption {
	return func(c *styleConfig) { c.fontSize = size }
}

// FontColor sets the font color as a hex string (e.g., "#FF0000" for red).
func FontColor(hex string) StyleOption {
	return func(c *styleConfig) { c.fontColor = hex }
}

// BgColor sets the cell background color as a hex string (e.g., "#4472C4").
func BgColor(hex string) StyleOption {
	return func(c *styleConfig) { c.bgColor = hex }
}

// AlignLeft aligns text to the left.
var AlignLeft StyleOption = func(c *styleConfig) { c.hAlign = "left" }

// AlignCenter centers text horizontally.
var AlignCenter StyleOption = func(c *styleConfig) { c.hAlign = "center" }

// AlignRight aligns text to the right.
var AlignRight StyleOption = func(c *styleConfig) { c.hAlign = "right" }

// VAlignTop aligns text to the top of the cell.
var VAlignTop StyleOption = func(c *styleConfig) { c.vAlign = "top" }

// VAlignMiddle centers text vertically.
var VAlignMiddle StyleOption = func(c *styleConfig) { c.vAlign = "center" }

// VAlignBottom aligns text to the bottom of the cell.
var VAlignBottom StyleOption = func(c *styleConfig) { c.vAlign = "bottom" }

// NumFmt sets a custom number format string (e.g., "#,##0.00", "0%", "yyyy-mm-dd").
func NumFmt(format string) StyleOption {
	return func(c *styleConfig) { c.numFmt = format }
}

// Border adds borders to cells. Specify the style as "thin", "medium", or "thick",
// and the color as a hex string.
func Border(style, color string) StyleOption {
	return func(c *styleConfig) {
		styleMap := map[string]int{
			"thin":   1,
			"medium": 2,
			"thick":  5,
		}
		s, ok := styleMap[style]
		if !ok {
			s = 1 // default thin
		}
		c.border = []excelize.Border{
			{Type: "left", Color: color, Style: s},
			{Type: "right", Color: color, Style: s},
			{Type: "top", Color: color, Style: s},
			{Type: "bottom", Color: color, Style: s},
		}
	}
}
