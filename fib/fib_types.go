package fib

// FileInformationBlock is the top-level structure for the FIB.
type FileInformationBlock struct {
	Base        FibBase
	Csw         uint16
	FibRgW      FibRgW97
	Cslw        uint16
	FibRgLw     FibRgLw97
	CbRgFcLcb   uint16
	RgFcLcbBlob []byte // Variable part, raw bytes for now
	// Parsed version for convenience
	RgFcLcb struct {
		FcClx  uint32
		LcbClx uint32
	}
	CswNew uint16
	// FibRgCswNew would follow here if present
}

// FibBase is the fixed-size (32 byte) header of the FIB.
type FibBase struct {
	WIdent   uint16
	NFib     uint16
	_        uint16 // unused
	Lid      uint16
	PnNext   uint16
	Flags1   uint16
	NFibBack uint16
	LKey     uint32
	Envr     byte
	Flags2   byte
	_        [2]uint16 // reserved
	_        [2]uint32 // reserved
}

// FibRgW97 is the 16-bit value section of the FIB.
type FibRgW97 struct {
	_ [28]byte // 14 values, stubbed for simplicity
}

// FibRgLw97 is the 32-bit value section of the FIB.
// We care about ccpText for text extraction.
type FibRgLw97 struct {
	_          [4]byte
	CcpText    uint32 // Count of characters in main document
	CcpFtn     uint32 // Count of characters in footnotes
	CcpHdd     uint32 // Count of characters in headers/footers
	_          [4]byte
	CcpAtn     uint32 // Count of characters in annotations
	CcpEdn     uint32 // Count of characters in endnotes
	CcpTxbx    uint32 // Count of characters in textboxes
	CcpHdrTxbx uint32 // Count of characters in header textboxes
	_          [44]byte
}
