package structures

import (
	"encoding/binary"
	"fmt"
)

// PCD (Piece Descriptor) describes a piece of text in the document.
// Each piece references a contiguous run of text in the WordDocument stream.
type PCD struct {
	FNoEncryption bool   // If true, piece is not encrypted
	FComplex      bool   // If true, piece contains complex formatting
	FC            uint32 // File Character position in WordDocument stream
	IsUnicode     bool   // If true, text is Unicode; if false, text is ANSI
}

// ParsePCD parses a PCD structure from an 8-byte data element.
func ParsePCD(data []byte) (*PCD, error) {
	if len(data) != 8 {
		return nil, fmt.Errorf("pcd: invalid data size %d, expected 8", len(data))
	}

	pcd := &PCD{}

	// First 2 bytes contain flags
	flags := binary.LittleEndian.Uint16(data[0:2])
	pcd.FNoEncryption = (flags & 0x0001) != 0
	pcd.FComplex = (flags & 0x0002) != 0

	// Next 4 bytes contain the file character position
	fc := binary.LittleEndian.Uint32(data[2:6])
	
	// Check if this is Unicode text
	pcd.IsUnicode = (fc & 0x40000000) != 0
	
	// Clear the Unicode flag to get the actual file position
	pcd.FC = fc & 0x3FFFFFFF

	return pcd, nil
}

// GetActualFC returns the actual file position for reading text.
// For Unicode text, the position needs to be divided by 2.
func (pcd *PCD) GetActualFC() uint32 {
	if pcd.IsUnicode {
		return pcd.FC / 2
	}
	return pcd.FC
}

// PlcPcd represents a PLC of Piece Descriptors (the piece table).
type PlcPcd struct {
	*PLC
	Pieces []*PCD
}

// ParsePlcPcd parses a piece table from raw data.
func ParsePlcPcd(data []byte) (*PlcPcd, error) {
	plc, err := ParsePLC(data, 8) // PCDs are 8 bytes each
	if err != nil {
		return nil, fmt.Errorf("plcpcd: failed to parse PLC: %w", err)
	}

	pieces := make([]*PCD, len(plc.Data))
	for i, pcdData := range plc.Data {
		pcd, err := ParsePCD(pcdData)
		if err != nil {
			return nil, fmt.Errorf("plcpcd: failed to parse PCD %d: %w", i, err)
		}
		pieces[i] = pcd
	}

	return &PlcPcd{
		PLC:    plc,
		Pieces: pieces,
	}, nil
}

// GetPieceAt returns the piece descriptor at the given index.
func (plcpcd *PlcPcd) GetPieceAt(index int) (*PCD, error) {
	if index < 0 || index >= len(plcpcd.Pieces) {
		return nil, fmt.Errorf("plcpcd: invalid index %d", index)
	}
	return plcpcd.Pieces[index], nil
}

// GetTextRange returns the character range and piece descriptor for a given piece index.
func (plcpcd *PlcPcd) GetTextRange(index int) (start, end CP, pcd *PCD, err error) {
	if index < 0 || index >= len(plcpcd.Pieces) {
		return 0, 0, nil, fmt.Errorf("plcpcd: invalid index %d", index)
	}
	
	start, end, err = plcpcd.GetRange(index)
	if err != nil {
		return 0, 0, nil, err
	}
	
	return start, end, plcpcd.Pieces[index], nil
}