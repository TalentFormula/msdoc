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
	RgFcLcb FibRgFcLcb97
	CswNew  uint16
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
	CbMac      uint32   // Size of main document text stream in bytes
	_          uint32   // reserved
	CcpText    uint32   // Count of characters in main document
	CcpFtn     uint32   // Count of characters in footnotes
	CcpHdd     uint32   // Count of characters in headers/footers
	_          uint32   // reserved
	CcpAtn     uint32   // Count of characters in annotations
	CcpEdn     uint32   // Count of characters in endnotes
	CcpTxbx    uint32   // Count of characters in textboxes
	CcpHdrTxbx uint32   // Count of characters in header textboxes
	_          [44]byte // remaining reserved fields
}

// FibRgFcLcb97 represents the file position and length pairs for Word 97 format.
// This structure contains pointers to various parts of the document.
type FibRgFcLcb97 struct {
	FcStshfOrig         uint32 // File position of original style sheet
	LcbStshfOrig        uint32 // Length of original style sheet
	FcStshf             uint32 // File position of style sheet
	LcbStshf            uint32 // Length of style sheet
	FcPlcffndRef        uint32 // File position of footnote reference PLC
	LcbPlcffndRef       uint32 // Length of footnote reference PLC
	FcPlcffndTxt        uint32 // File position of footnote text PLC
	LcbPlcffndTxt       uint32 // Length of footnote text PLC
	FcPlcfandRef        uint32 // File position of annotation reference PLC
	LcbPlcfandRef       uint32 // Length of annotation reference PLC
	FcPlcfandTxt        uint32 // File position of annotation text PLC
	LcbPlcfandTxt       uint32 // Length of annotation text PLC
	FcPlcfsed           uint32 // File position of section descriptor PLC
	LcbPlcfsed          uint32 // Length of section descriptor PLC
	FcPlcfpgdFtn        uint32 // File position of page descriptor PLC for footnotes
	LcbPlcfpgdFtn       uint32 // Length of page descriptor PLC for footnotes
	FcPlcfhdd           uint32 // File position of header document PLC
	LcbPlcfhdd          uint32 // Length of header document PLC
	FcPlcfbteChpx       uint32 // File position of character property bin table PLC
	LcbPlcfbteChpx      uint32 // Length of character property bin table PLC
	FcPlcfbtePapx       uint32 // File position of paragraph property bin table PLC
	LcbPlcfbtePapx      uint32 // Length of paragraph property bin table PLC
	FcPlcfsea           uint32 // File position of private use PLC
	LcbPlcfsea          uint32 // Length of private use PLC
	FcSttbfffn          uint32 // File position of font information STTB
	LcbSttbfffn         uint32 // Length of font information STTB
	FcPlcffldMom        uint32 // File position of field PLC for main document
	LcbPlcffldMom       uint32 // Length of field PLC for main document
	FcPlcffldHdr        uint32 // File position of field PLC for header document
	LcbPlcffldHdr       uint32 // Length of field PLC for header document
	FcPlcffldFtn        uint32 // File position of field PLC for footnote document
	LcbPlcffldFtn       uint32 // Length of field PLC for footnote document
	FcPlcffldAtn        uint32 // File position of field PLC for annotation document
	LcbPlcffldAtn       uint32 // Length of field PLC for annotation document
	FcPlcffldMcr        uint32 // File position of field PLC for macro document
	LcbPlcffldMcr       uint32 // Length of field PLC for macro document
	FcSttbfbkmk         uint32 // File position of bookmark names STTB
	LcbSttbfbkmk        uint32 // Length of bookmark names STTB
	FcPlcfbkf           uint32 // File position of bookmark start PLC
	LcbPlcfbkf          uint32 // Length of bookmark start PLC
	FcPlcfbkl           uint32 // File position of bookmark end PLC
	LcbPlcfbkl          uint32 // Length of bookmark end PLC
	FcCmds              uint32 // File position of macro commands
	LcbCmds             uint32 // Length of macro commands
	FcPlcfmcr           uint32 // File position of macro PLC
	LcbPlcfmcr          uint32 // Length of macro PLC
	FcSttbfmcr          uint32 // File position of macro names STTB
	LcbSttbfmcr         uint32 // Length of macro names STTB
	FcPrDrvr            uint32 // File position of printer driver information
	LcbPrDrvr           uint32 // Length of printer driver information
	FcPrEnvPort         uint32 // File position of print environment in portrait mode
	LcbPrEnvPort        uint32 // Length of print environment in portrait mode
	FcPrEnvLand         uint32 // File position of print environment in landscape mode
	LcbPrEnvLand        uint32 // Length of print environment in landscape mode
	FcWss               uint32 // File position of window save state
	LcbWss              uint32 // Length of window save state
	FcDop               uint32 // File position of document properties
	LcbDop              uint32 // Length of document properties
	FcSttbfAssoc        uint32 // File position of associated strings STTB
	LcbSttbfAssoc       uint32 // Length of associated strings STTB
	FcClx               uint32 // File position of character and paragraph formatting PLC
	LcbClx              uint32 // Length of character and paragraph formatting PLC
	FcPlcfpgdFtn2       uint32 // File position of page descriptor PLC for footnotes
	LcbPlcfpgdFtn2      uint32 // Length of page descriptor PLC for footnotes
	FcPlcfpgdEdn        uint32 // File position of page descriptor PLC for endnotes
	LcbPlcfpgdEdn       uint32 // Length of page descriptor PLC for endnotes
	FcPlcfpgdEdn2       uint32 // File position of page descriptor PLC for endnotes
	LcbPlcfpgdEdn2      uint32 // Length of page descriptor PLC for endnotes
	FcDggInfo           uint32 // File position of drawing objects
	LcbDggInfo          uint32 // Length of drawing objects
	FcSttbfRMark        uint32 // File position of revision mark authors STTB
	LcbSttbfRMark       uint32 // Length of revision mark authors STTB
	FcSttbfCaption      uint32 // File position of caption STTB
	LcbSttbfCaption     uint32 // Length of caption STTB
	FcSttbfAutoCaption  uint32 // File position of auto caption STTB
	LcbSttbfAutoCaption uint32 // Length of auto caption STTB
	FcPlcfwkb           uint32 // File position of WKB PLC
	LcbPlcfwkb          uint32 // Length of WKB PLC
	FcPlcfspl           uint32 // File position of spell check state PLC
	LcbPlcfspl          uint32 // Length of spell check state PLC
	FcPlcftxbxTxt       uint32 // File position of textbox break table PLC
	LcbPlcftxbxTxt      uint32 // Length of textbox break table PLC
	FcPlcffldTxbx       uint32 // File position of textbox field table PLC
	LcbPlcffldTxbx      uint32 // Length of textbox field table PLC
	FcPlcfhdrtxbxTxt    uint32 // File position of header textbox break table PLC
	LcbPlcfhdrtxbxTxt   uint32 // Length of header textbox break table PLC
	FcPlcffldHdrTxbx    uint32 // File position of header textbox field table PLC
	LcbPlcffldHdrTxbx   uint32 // Length of header textbox field table PLC
	FcStwUser           uint32 // File position of user-defined table
	LcbStwUser          uint32 // Length of user-defined table
	FcSttbttmbd         uint32 // File position of embedded TrueType font data
	LcbSttbttmbd        uint32 // Length of embedded TrueType font data
	// Additional fields would continue for different nFib versions...
}
