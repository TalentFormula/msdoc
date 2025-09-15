package streams

import (
	"fmt"

	"github.com/TalentFormula/msdoc/structures"
)

// TableStream represents either the 0Table or 1Table stream containing formatting information.
type TableStream struct {
	Data []byte
	Name string // "0Table" or "1Table"
}

// NewTableStream creates a new Table stream processor.
func NewTableStream(data []byte, name string) *TableStream {
	return &TableStream{
		Data: data,
		Name: name,
	}
}

// GetPieceTable extracts the piece table (PlcPcd) from the specified location.
func (ts *TableStream) GetPieceTable(fcClx, lcbClx uint32) (*structures.PlcPcd, error) {
	if lcbClx == 0 {
		return nil, fmt.Errorf("table: no piece table data")
	}

	if fcClx+lcbClx > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: piece table location out of bounds")
	}

	clx := ts.Data[fcClx : fcClx+lcbClx]

	// The CLX should start with a PlcPcd marker (0x02)
	if len(clx) == 0 || clx[0] != 0x02 {
		return nil, fmt.Errorf("table: invalid CLX structure, expected PlcPcd marker")
	}

	// Parse the piece table
	plcPcdData := clx[1:] // Skip the marker byte
	return structures.ParsePlcPcd(plcPcdData)
}

// GetStyleSheet extracts the style sheet from the specified location.
func (ts *TableStream) GetStyleSheet(fcStsh, lcbStsh uint32) ([]byte, error) {
	if lcbStsh == 0 {
		return nil, nil // No style sheet
	}

	if fcStsh+lcbStsh > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: style sheet location out of bounds")
	}

	result := make([]byte, lcbStsh)
	copy(result, ts.Data[fcStsh:fcStsh+lcbStsh])
	return result, nil
}

// GetSectionTable extracts the section descriptor PLC from the specified location.
func (ts *TableStream) GetSectionTable(fcPlcfsed, lcbPlcfsed uint32) (*structures.PLC, error) {
	if lcbPlcfsed == 0 {
		return nil, nil // No section table
	}

	if fcPlcfsed+lcbPlcfsed > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: section table location out of bounds")
	}

	sedData := ts.Data[fcPlcfsed : fcPlcfsed+lcbPlcfsed]
	// Section descriptors are 12 bytes each
	return structures.ParsePLC(sedData, 12)
}

// GetCharacterFormattingTable extracts the character formatting FKP references.
func (ts *TableStream) GetCharacterFormattingTable(fcPlcfbteChpx, lcbPlcfbteChpx uint32) (*structures.PLC, error) {
	if lcbPlcfbteChpx == 0 {
		return nil, nil // No character formatting table
	}

	if fcPlcfbteChpx+lcbPlcfbteChpx > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: character formatting table location out of bounds")
	}

	chpxData := ts.Data[fcPlcfbteChpx : fcPlcfbteChpx+lcbPlcfbteChpx]
	// BTE (Bin Table Entry) structures are 4 bytes each
	return structures.ParsePLC(chpxData, 4)
}

// GetParagraphFormattingTable extracts the paragraph formatting FKP references.
func (ts *TableStream) GetParagraphFormattingTable(fcPlcfbtePapx, lcbPlcfbtePapx uint32) (*structures.PLC, error) {
	if lcbPlcfbtePapx == 0 {
		return nil, nil // No paragraph formatting table
	}

	if fcPlcfbtePapx+lcbPlcfbtePapx > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: paragraph formatting table location out of bounds")
	}

	papxData := ts.Data[fcPlcfbtePapx : fcPlcfbtePapx+lcbPlcfbtePapx]
	// BTE (Bin Table Entry) structures are 4 bytes each
	return structures.ParsePLC(papxData, 4)
}

// GetFontTable extracts the font information table from the specified location.
func (ts *TableStream) GetFontTable(fcSttbfffn, lcbSttbfffn uint32) ([]byte, error) {
	if lcbSttbfffn == 0 {
		return nil, nil // No font table
	}

	if fcSttbfffn+lcbSttbfffn > uint32(len(ts.Data)) {
		return nil, fmt.Errorf("table: font table location out of bounds")
	}

	result := make([]byte, lcbSttbfffn)
	copy(result, ts.Data[fcSttbfffn:fcSttbfffn+lcbSttbfffn])
	return result, nil
}

// IsEncrypted checks if this table stream contains encryption information.
func (ts *TableStream) IsEncrypted() bool {
	// For encrypted documents, the table stream starts with an EncryptionHeader
	// This is a simplified check - a complete implementation would parse the header
	return len(ts.Data) > 4 && ts.Data[0] == 0x01 // Basic check for encryption marker
}

// This file would contain logic for parsing the 0Table or 1Table streams,
// such as character and paragraph formatting tables.
