// Package formatting provides comprehensive support for text formatting in .doc files.
//
// This package exposes character properties (CHPX), paragraph properties (PAPX),
// and section properties (SEP) with full formatting information including fonts,
// colors, styles, and layout properties.
package formatting

import (
	"encoding/binary"
	"fmt"
)

// TextRun represents a run of text with consistent formatting.
type TextRun struct {
	Text      string              // The actual text content
	StartPos  uint32              // Starting character position
	EndPos    uint32              // Ending character position  
	CharProps *CharacterProperties // Character formatting properties
	ParaProps *ParagraphProperties // Paragraph formatting properties (if paragraph boundary)
}

// CharacterProperties holds all character-level formatting information.
type CharacterProperties struct {
	FontName       string      // Font family name
	FontSize       uint16      // Font size in half-points (e.g., 24 = 12pt)
	Bold           bool        // Bold formatting
	Italic         bool        // Italic formatting
	Underline      UnderlineType // Underline style
	Strikethrough  bool        // Strikethrough formatting
	Superscript    bool        // Superscript formatting
	Subscript      bool        // Subscript formatting
	Color          Color       // Text color
	HighlightColor Color       // Highlight/background color
	FontCharset    uint8       // Character set (for non-ASCII text)
	Language       uint16      // Language identifier
	Hidden         bool        // Hidden text
	SmallCaps      bool        // Small capitals
	AllCaps        bool        // All capitals
	Spacing        int16       // Character spacing in twips
	Scale          uint16      // Horizontal scaling percentage
	Position       int16       // Vertical position offset
	Border         *Border     // Character border
	Shading        *Shading    // Character shading
}

// ParagraphProperties holds all paragraph-level formatting information.
type ParagraphProperties struct {
	Alignment      ParagraphAlignment // Text alignment
	LeftIndent     int32              // Left indent in twips
	RightIndent    int32              // Right indent in twips
	FirstLineIndent int32             // First line indent in twips
	SpaceBefore    uint16             // Space before paragraph in twips
	SpaceAfter     uint16             // Space after paragraph in twips
	LineSpacing    LineSpacing        // Line spacing configuration
	KeepWithNext   bool               // Keep with next paragraph
	KeepTogether   bool               // Keep paragraph together
	PageBreakBefore bool              // Force page break before
	WidowControl   bool               // Widow/orphan control
	Borders        *ParagraphBorders  // Paragraph borders
	Shading        *Shading           // Paragraph shading
	TabStops       []TabStop          // Tab stop positions
	OutlineLevel   uint8              // Outline level (0-9)
	StyleName      string             // Applied paragraph style name
}

// SectionProperties holds section-level formatting information.
type SectionProperties struct {
	PageWidth      uint32            // Page width in twips
	PageHeight     uint32            // Page height in twips
	LeftMargin     uint32            // Left margin in twips
	RightMargin    uint32            // Right margin in twips
	TopMargin      uint32            // Top margin in twips
	BottomMargin   uint32            // Bottom margin in twips
	HeaderMargin   uint32            // Header margin in twips
	FooterMargin   uint32            // Footer margin in twips
	Orientation    PageOrientation   // Page orientation
	Columns        uint16            // Number of columns
	ColumnSpacing  uint32            // Column spacing in twips
	VerticalAlign  VerticalAlignment // Vertical alignment
	PageBorders    *PageBorders      // Page borders
	PageBackground *Shading          // Page background
	LineNumbers    *LineNumbering    // Line numbering settings
	EndnoteProps   *EndnoteProperties // Endnote properties
	FootnoteProps  *FootnoteProperties // Footnote properties
}

// UnderlineType represents different underline styles.
type UnderlineType int

const (
	UnderlineNone UnderlineType = iota
	UnderlineSingle
	UnderlineDouble
	UnderlineDotted
	UnderlineDashed
	UnderlineWavy
	UnderlineThick
)

// Color represents an RGB color value.
type Color struct {
	Red   uint8 // Red component (0-255)
	Green uint8 // Green component (0-255)  
	Blue  uint8 // Blue component (0-255)
	Auto  bool  // True if color is automatic
}

// ParagraphAlignment represents text alignment options.
type ParagraphAlignment int

const (
	AlignLeft ParagraphAlignment = iota
	AlignCenter
	AlignRight
	AlignJustify
	AlignDistribute
)

// PageOrientation represents page orientation.
type PageOrientation int

const (
	OrientationPortrait PageOrientation = iota
	OrientationLandscape
)

// VerticalAlignment represents vertical text alignment on page.
type VerticalAlignment int

const (
	VerticalAlignTop VerticalAlignment = iota
	VerticalAlignCenter
	VerticalAlignBottom
	VerticalAlignJustify
)

// LineSpacing represents line spacing configuration.
type LineSpacing struct {
	Type  LineSpacingType // Spacing type
	Value uint16          // Spacing value (interpretation depends on type)
}

// LineSpacingType represents different line spacing modes.
type LineSpacingType int

const (
	LineSpacingSingle LineSpacingType = iota
	LineSpacingOneAndHalf
	LineSpacingDouble
	LineSpacingAtLeast
	LineSpacingExact
	LineSpacingMultiple
)

// Border represents border formatting.
type Border struct {
	Style     BorderStyle // Border style
	Width     uint16      // Border width in eighth-points
	Color     Color       // Border color
	Spacing   uint16      // Distance from text in points
	Shadow    bool        // Drop shadow effect
}

// BorderStyle represents different border styles.
type BorderStyle int

const (
	BorderNone BorderStyle = iota
	BorderSingle
	BorderDouble
	BorderDotted
	BorderDashed
	BorderThick
	BorderThinThickSmall
	BorderThickThinSmall
)

// Shading represents background shading/fill.
type Shading struct {
	Pattern     ShadingPattern // Fill pattern
	ForeColor   Color          // Foreground color
	BackColor   Color          // Background color
	Percentage  uint16         // Fill percentage (0-100)
}

// ShadingPattern represents fill patterns.
type ShadingPattern int

const (
	ShadingClear ShadingPattern = iota
	ShadingSolid
	ShadingPct5
	ShadingPct10
	ShadingPct20
	ShadingPct25
	ShadingPct30
	ShadingPct40
	ShadingPct50
	ShadingPct60
	ShadingPct70
	ShadingPct75
	ShadingPct80
	ShadingPct90
)

// ParagraphBorders represents borders around a paragraph.
type ParagraphBorders struct {
	Top    *Border // Top border
	Left   *Border // Left border
	Bottom *Border // Bottom border
	Right  *Border // Right border
	Box    *Border // Box border (all sides)
	Bar    *Border // Bar border (left side)
}

// PageBorders represents borders around a page.
type PageBorders struct {
	Top    *Border // Top border
	Left   *Border // Left border  
	Bottom *Border // Bottom border
	Right  *Border // Right border
}

// TabStop represents a tab stop position.
type TabStop struct {
	Position  uint32        // Tab position in twips
	Type      TabStopType   // Tab type
	Leader    TabLeader     // Tab leader character
}

// TabStopType represents different tab stop types.
type TabStopType int

const (
	TabLeft TabStopType = iota
	TabCenter
	TabRight
	TabDecimal
	TabBar
)

// TabLeader represents tab leader characters.
type TabLeader int

const (
	TabLeaderNone TabLeader = iota
	TabLeaderDots
	TabLeaderDashes
	TabLeaderUnderline
	TabLeaderThickLine
)

// LineNumbering represents line numbering settings.
type LineNumbering struct {
	Start    uint16 // Starting line number
	Distance uint32 // Distance from text in twips
	Interval uint16 // Numbering interval (every Nth line)
	Restart  LineNumberRestart // Restart behavior
}

// LineNumberRestart represents line number restart options.
type LineNumberRestart int

const (
	LineNumberContinuous LineNumberRestart = iota
	LineNumberRestartPage
	LineNumberRestartSection
)

// EndnoteProperties represents endnote formatting settings.
type EndnoteProperties struct {
	NumberFormat EndnoteNumberFormat // Number format
	StartNumber  uint16              // Starting number
	Restart      EndnoteRestart      // Restart behavior
	Position     EndnotePosition     // Position in document
}

// FootnoteProperties represents footnote formatting settings.
type FootnoteProperties struct {
	NumberFormat FootnoteNumberFormat // Number format
	StartNumber  uint16               // Starting number
	Restart      FootnoteRestart      // Restart behavior
	Position     FootnotePosition     // Position on page
}

// EndnoteNumberFormat represents endnote numbering formats.
type EndnoteNumberFormat int

const (
	EndnoteArabic EndnoteNumberFormat = iota
	EndnoteRomanUpper
	EndnoteRomanLower
	EndnoteLetterUpper
	EndnoteLetterLower
)

// FootnoteNumberFormat represents footnote numbering formats.
type FootnoteNumberFormat int

const (
	FootnoteArabic FootnoteNumberFormat = iota
	FootnoteRomanUpper
	FootnoteRomanLower
	FootnoteLetterUpper
	FootnoteLetterLower
	FootnoteSymbol
)

// EndnoteRestart represents endnote restart options.
type EndnoteRestart int

const (
	EndnoteContinuous EndnoteRestart = iota
	EndnoteRestartSection
)

// FootnoteRestart represents footnote restart options.
type FootnoteRestart int

const (
	FootnoteContinuous FootnoteRestart = iota
	FootnoteRestartPage
	FootnoteRestartSection
)

// EndnotePosition represents endnote position options.
type EndnotePosition int

const (
	EndnoteEndOfSection EndnotePosition = iota
	EndnoteEndOfDocument
)

// FootnotePosition represents footnote position options.
type FootnotePosition int

const (
	FootnoteBottomOfPage FootnotePosition = iota
	FootnoteBelowText
)

// FormattingExtractor extracts formatting information from FKP structures.
type FormattingExtractor struct {
	fontTable map[uint16]string // Font table mapping
	styleTable map[uint16]string // Style table mapping
}

// NewFormattingExtractor creates a new formatting extractor.
func NewFormattingExtractor() *FormattingExtractor {
	return &FormattingExtractor{
		fontTable:  make(map[uint16]string),
		styleTable: make(map[uint16]string),
	}
}

// ParseCharacterProperties parses character properties from CHPX data.
func (fe *FormattingExtractor) ParseCharacterProperties(chpx []byte) (*CharacterProperties, error) {
	if len(chpx) < 2 {
		return nil, fmt.Errorf("CHPX data too short")
	}

	props := &CharacterProperties{
		FontSize: 24, // Default 12pt
		Color:    Color{Auto: true},
		Scale:    100, // Default 100%
	}

	// Parse CHPX properties
	offset := 0
	for offset < len(chpx)-1 {
		sprm := binary.LittleEndian.Uint16(chpx[offset:])
		offset += 2

		switch sprm {
		case 0x4A03: // Font size
			if offset < len(chpx) {
				props.FontSize = uint16(chpx[offset]) * 2 // Convert to half-points
				offset++
			}
		case 0x085C: // Bold
			if offset < len(chpx) {
				props.Bold = chpx[offset] != 0
				offset++
			}
		case 0x085D: // Italic
			if offset < len(chpx) {
				props.Italic = chpx[offset] != 0
				offset++
			}
		case 0x2A0C: // Font color
			if offset+2 < len(chpx) {
				colorVal := binary.LittleEndian.Uint16(chpx[offset:])
				props.Color = fe.parseColor(colorVal)
				offset += 2
			}
		default:
			// Skip unknown properties
			offset++
		}
	}

	return props, nil
}

// ParseParagraphProperties parses paragraph properties from PAPX data.
func (fe *FormattingExtractor) ParseParagraphProperties(papx []byte) (*ParagraphProperties, error) {
	if len(papx) < 2 {
		return nil, fmt.Errorf("PAPX data too short")
	}

	props := &ParagraphProperties{
		Alignment:   AlignLeft,
		LineSpacing: LineSpacing{Type: LineSpacingSingle, Value: 240}, // Default single spacing
	}

	// Parse PAPX properties
	offset := 0
	for offset < len(papx)-1 {
		sprm := binary.LittleEndian.Uint16(papx[offset:])
		offset += 2

		switch sprm {
		case 0x2405: // Paragraph alignment
			if offset < len(papx) {
				props.Alignment = ParagraphAlignment(papx[offset])
				offset++
			}
		case 0x840E: // Left indent
			if offset+2 < len(papx) {
				props.LeftIndent = int32(binary.LittleEndian.Uint16(papx[offset:]))
				offset += 2
			}
		case 0x8411: // Right indent
			if offset+2 < len(papx) {
				props.RightIndent = int32(binary.LittleEndian.Uint16(papx[offset:]))
				offset += 2
			}
		default:
			// Skip unknown properties
			offset++
		}
	}

	return props, nil
}

// parseColor converts a Word color value to a Color struct.
func (fe *FormattingExtractor) parseColor(colorVal uint16) Color {
	if colorVal == 0xFF {
		return Color{Auto: true}
	}

	// Standard Word color palette
	colors := []Color{
		{0, 0, 0, false},       // Black
		{0, 0, 255, false},     // Blue
		{0, 255, 255, false},   // Cyan
		{0, 255, 0, false},     // Green
		{255, 0, 255, false},   // Magenta
		{255, 0, 0, false},     // Red
		{255, 255, 0, false},   // Yellow
		{255, 255, 255, false}, // White
	}

	if int(colorVal) < len(colors) {
		return colors[colorVal]
	}

	// Custom color - extract RGB components
	return Color{
		Red:   uint8((colorVal >> 0) & 0xFF),
		Green: uint8((colorVal >> 8) & 0xFF),
		Blue:  uint8((colorVal >> 16) & 0xFF),
		Auto:  false,
	}
}

// AddFontMapping adds a font mapping to the font table.
func (fe *FormattingExtractor) AddFontMapping(fontID uint16, fontName string) {
	fe.fontTable[fontID] = fontName
}

// AddStyleMapping adds a style mapping to the style table.
func (fe *FormattingExtractor) AddStyleMapping(styleID uint16, styleName string) {
	fe.styleTable[styleID] = styleName
}