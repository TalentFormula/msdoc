package structures

import (
	"encoding/binary"
	"fmt"
)

// SEPX (Section Property eXtension) contains section-level formatting information.
type SEPX struct {
	Data   []byte // Raw SEPX data
	Length uint16 // Length of the SEPX data
}

// SEP (Section Properties) contains parsed section formatting information.
type SEP struct {
	// Page setup
	XaPage       uint16 // Page width in twips
	YaPage       uint16 // Page height in twips
	DxaLeft      uint16 // Left margin in twips
	DxaRight     uint16 // Right margin in twips
	DyaTop       uint16 // Top margin in twips
	DyaBottom    uint16 // Bottom margin in twips
	DyaHdrTop    uint16 // Header top margin in twips
	DyaHdrBottom uint16 // Header bottom margin in twips

	// Page orientation and layout
	FLandscape  bool   // True if landscape orientation
	FContinuous bool   // True if continuous section break
	FTitlePage  bool   // True if different first page
	FPgnRestart bool   // True if restart page numbering
	PgnStart    uint16 // Starting page number

	// Column layout
	CcolM1        uint16 // Number of columns minus 1
	FEvenlySpaced bool   // True if columns are evenly spaced
	DxaColumns    uint16 // Space between columns in twips

	// Line numbering
	Lnc    uint8  // Line number count
	DxaLnn uint16 // Distance from text to line numbers
	LnnMin uint16 // Starting line number

	// Headers and footers
	GrpfIhdt uint8 // Header/footer flags
}

// ParseSEPX parses a SEPX structure from raw data.
func ParseSEPX(data []byte) (*SEPX, error) {
	if len(data) < 2 {
		return nil, fmt.Errorf("sepx: data too short for length field")
	}

	length := binary.LittleEndian.Uint16(data[0:2])
	if int(length) > len(data)-2 {
		return nil, fmt.Errorf("sepx: invalid length %d for data size %d", length, len(data))
	}

	sepx := &SEPX{
		Length: length,
		Data:   make([]byte, length),
	}

	if length > 0 {
		copy(sepx.Data, data[2:2+length])
	}

	return sepx, nil
}

// ParseSEP parses a SEP structure from SEPX data.
func (sepx *SEPX) ParseSEP() (*SEP, error) {
	if len(sepx.Data) < 48 { // Minimum size for basic SEP fields
		return nil, fmt.Errorf("sepx: data too short for SEP structure")
	}

	sep := &SEP{}
	data := sepx.Data

	// Parse basic page layout fields
	sep.XaPage = binary.LittleEndian.Uint16(data[0:2])
	sep.YaPage = binary.LittleEndian.Uint16(data[2:4])
	sep.DxaLeft = binary.LittleEndian.Uint16(data[4:6])
	sep.DxaRight = binary.LittleEndian.Uint16(data[6:8])
	sep.DyaTop = binary.LittleEndian.Uint16(data[8:10])
	sep.DyaBottom = binary.LittleEndian.Uint16(data[10:12])
	sep.DyaHdrTop = binary.LittleEndian.Uint16(data[12:14])
	sep.DyaHdrBottom = binary.LittleEndian.Uint16(data[14:16])

	// Parse flags from various positions
	if len(data) > 16 {
		flags := binary.LittleEndian.Uint16(data[16:18])
		sep.FLandscape = (flags & 0x0001) != 0
		sep.FContinuous = (flags & 0x0002) != 0
		sep.FTitlePage = (flags & 0x0004) != 0
		sep.FPgnRestart = (flags & 0x0008) != 0
	}

	if len(data) > 18 {
		sep.PgnStart = binary.LittleEndian.Uint16(data[18:20])
	}

	// Parse column information
	if len(data) > 20 {
		sep.CcolM1 = binary.LittleEndian.Uint16(data[20:22])
		colFlags := binary.LittleEndian.Uint16(data[22:24])
		sep.FEvenlySpaced = (colFlags & 0x0001) != 0
		sep.DxaColumns = binary.LittleEndian.Uint16(data[24:26])
	}

	// Parse line numbering
	if len(data) > 26 {
		sep.Lnc = data[26]
		sep.DxaLnn = binary.LittleEndian.Uint16(data[27:29])
		sep.LnnMin = binary.LittleEndian.Uint16(data[29:31])
	}

	// Parse header/footer information
	if len(data) > 31 {
		sep.GrpfIhdt = data[31]
	}

	return sep, nil
}

// IsLandscape returns true if the section uses landscape orientation.
func (sep *SEP) IsLandscape() bool {
	return sep.FLandscape
}

// GetPageDimensions returns the page width and height in twips.
func (sep *SEP) GetPageDimensions() (width, height uint16) {
	return sep.XaPage, sep.YaPage
}

// GetMargins returns the page margins in twips (left, right, top, bottom).
func (sep *SEP) GetMargins() (left, right, top, bottom uint16) {
	return sep.DxaLeft, sep.DxaRight, sep.DyaTop, sep.DyaBottom
}

// GetColumnCount returns the number of columns in the section.
func (sep *SEP) GetColumnCount() int {
	return int(sep.CcolM1) + 1
}

// HasDifferentFirstPage returns true if the section has a different first page.
func (sep *SEP) HasDifferentFirstPage() bool {
	return sep.FTitlePage
}

// SEP (Section Properties) defines formatting for a document section,
// such as page size, margins, and column layout.
