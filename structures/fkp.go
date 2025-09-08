package structures

import (
	"encoding/binary"
	"fmt"
)

// FKP (Formatted Disk Page) is a 512-byte page containing formatting data.
// There are different types of FKPs for different kinds of formatting:
// - CHPX FKP: Character properties (formatting like bold, italic, etc.)
// - PAPX FKP: Paragraph properties (formatting like alignment, spacing, etc.)
const FKPSize = 512

// FKP represents a formatted disk page containing formatting information.
type FKP struct {
	Data       []byte // Raw 512-byte page data
	Type       FKPType
	Entries    []FKPEntry
	EntryCount int
}

// FKPType indicates the type of formatting stored in the FKP.
type FKPType int

const (
	FKPTypeUnknown FKPType = iota
	FKPTypeCHP             // Character properties
	FKPTypePAP             // Paragraph properties
)

// FKPEntry represents a single formatting entry within an FKP.
type FKPEntry struct {
	FC     uint32 // File character position
	Offset uint16 // Offset within the FKP to the formatting data
	Data   []byte // The actual formatting data
}

// ParseFKP parses an FKP from raw 512-byte page data.
func ParseFKP(data []byte, fkpType FKPType) (*FKP, error) {
	if len(data) != FKPSize {
		return nil, fmt.Errorf("fkp: invalid data size %d, expected %d", len(data), FKPSize)
	}

	fkp := &FKP{
		Data: make([]byte, FKPSize),
		Type: fkpType,
	}
	copy(fkp.Data, data)

	// The last byte contains the count of entries
	entryCount := int(data[FKPSize-1])
	if entryCount > 255 {
		return nil, fmt.Errorf("fkp: invalid entry count %d", entryCount)
	}
	fkp.EntryCount = entryCount

	// Parse entries based on FKP type
	switch fkpType {
	case FKPTypeCHP:
		return parseCHPXFKP(fkp)
	case FKPTypePAP:
		return parsePAPXFKP(fkp)
	default:
		return fkp, nil // Return basic FKP without parsing entries
	}
}

// parseCHPXFKP parses a character properties FKP.
func parseCHPXFKP(fkp *FKP) (*FKP, error) {
	entryCount := fkp.EntryCount
	
	// Validate that we have enough space for the entries
	// Each entry is 5 bytes (4 bytes FC + 1 byte offset)
	if entryCount*5 > FKPSize-1 { // -1 for the count byte
		return nil, fmt.Errorf("fkp: too many entries (%d) for CHPX FKP", entryCount)
	}
	
	entries := make([]FKPEntry, entryCount)

	// Each entry consists of a 4-byte FC followed by a 1-byte offset
	for i := 0; i < entryCount; i++ {
		entryOffset := i * 5 // 4 bytes FC + 1 byte offset
		if entryOffset+5 > len(fkp.Data) {
			return nil, fmt.Errorf("fkp: entry %d out of bounds", i)
		}

		fc := binary.LittleEndian.Uint32(fkp.Data[entryOffset : entryOffset+4])
		offset := uint16(fkp.Data[entryOffset+4])

		entry := FKPEntry{
			FC:     fc,
			Offset: offset,
		}

		// Extract the actual formatting data if offset is valid
		if offset > 0 && int(offset) < FKPSize {
			// The formatting data starts at the offset
			// For CHPX, the first byte indicates the length
			if int(offset) < len(fkp.Data) {
				length := int(fkp.Data[offset])
				endPos := int(offset) + 1 + length
				if length > 0 && endPos <= len(fkp.Data) {
					entry.Data = make([]byte, length)
					copy(entry.Data, fkp.Data[int(offset)+1:endPos])
				}
			}
		}

		entries[i] = entry
	}

	fkp.Entries = entries
	return fkp, nil
}

// parsePAPXFKP parses a paragraph properties FKP.
func parsePAPXFKP(fkp *FKP) (*FKP, error) {
	entryCount := fkp.EntryCount
	
	// Validate that we have enough space for the entries
	// Each entry is 6 bytes (4 bytes FC + 2 bytes offset)
	if entryCount*6 > FKPSize-1 { // -1 for the count byte
		return nil, fmt.Errorf("fkp: too many entries (%d) for PAPX FKP", entryCount)
	}
	
	entries := make([]FKPEntry, entryCount)

	// Each entry consists of a 4-byte FC followed by a 2-byte offset
	for i := 0; i < entryCount; i++ {
		entryOffset := i * 6 // 4 bytes FC + 2 bytes offset
		if entryOffset+6 > len(fkp.Data) {
			return nil, fmt.Errorf("fkp: entry %d out of bounds", i)
		}

		fc := binary.LittleEndian.Uint32(fkp.Data[entryOffset : entryOffset+4])
		offset := binary.LittleEndian.Uint16(fkp.Data[entryOffset+4 : entryOffset+6])

		entry := FKPEntry{
			FC:     fc,
			Offset: offset,
		}

		// Extract the actual formatting data if offset is valid
		if offset > 0 && int(offset) < FKPSize {
			// For PAPX, the first byte indicates the length (multiply by 2)
			if int(offset) < len(fkp.Data) {
				lengthWords := int(fkp.Data[offset])
				length := lengthWords * 2
				endPos := int(offset) + 1 + length
				if length > 0 && endPos <= len(fkp.Data) {
					entry.Data = make([]byte, length)
					copy(entry.Data, fkp.Data[int(offset)+1:endPos])
				}
			}
		}

		entries[i] = entry
	}

	fkp.Entries = entries
	return fkp, nil
}

// GetEntryAt returns the formatting entry at the given index.
func (fkp *FKP) GetEntryAt(index int) (*FKPEntry, error) {
	if index < 0 || index >= len(fkp.Entries) {
		return nil, fmt.Errorf("fkp: invalid entry index %d", index)
	}
	return &fkp.Entries[index], nil
}

// FindEntryForFC finds the formatting entry that applies to the given file character position.
func (fkp *FKP) FindEntryForFC(fc uint32) *FKPEntry {
	// Find the entry with the highest FC that is <= the target FC
	var bestEntry *FKPEntry
	for i := range fkp.Entries {
		entry := &fkp.Entries[i]
		if entry.FC <= fc && (bestEntry == nil || entry.FC > bestEntry.FC) {
			bestEntry = entry
		}
	}
	return bestEntry
}
