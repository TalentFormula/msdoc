package fib

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
)

// ParseFIB reads a byte slice (from the WordDocument stream)
// and parses it into a FileInformationBlock struct.
func ParseFIB(data []byte) (*FileInformationBlock, error) {
	if len(data) < 32 { // Minimum size for FibBase
		return nil, errors.New("fib: data too short for FibBase")
	}

	r := bytes.NewReader(data)
	fib := &FileInformationBlock{}

	// Read the fixed-size FibBase
	if err := binary.Read(r, binary.LittleEndian, &fib.Base); err != nil {
		return nil, fmt.Errorf("fib: could not read FibBase: %w", err)
	}

	// Validate Word document identifier
	if fib.Base.WIdent != 0xA5EC {
		return nil, errors.New("fib: invalid wIdent, not a Word document")
	}

	// Read remaining FIB sections
	currentOffset, _ := r.Seek(0, 1) // Get current position
	if err := binary.Read(r, binary.LittleEndian, &fib.Csw); err != nil {
		return nil, fmt.Errorf("fib: failed to read Csw at offset %d: %w", currentOffset, err)
	}

	// Skip FibRgW
	fibRgWBytes := make([]byte, 28)
	if _, err := r.Read(fibRgWBytes); err != nil {
		return nil, fmt.Errorf("fib: failed to read FibRgW: %w", err)
	}

	currentOffset, _ = r.Seek(0, 1)
	if err := binary.Read(r, binary.LittleEndian, &fib.Cslw); err != nil {
		return nil, fmt.Errorf("fib: failed to read Cslw at offset %d: %w", currentOffset, err)
	}

	// Skip FibRgLw
	fibRgLwBytes := make([]byte, 76) // Known size for FibRgLw97
	if _, err := r.Read(fibRgLwBytes); err != nil {
		return nil, fmt.Errorf("fib: failed to read FibRgLw: %w", err)
	}

	currentOffset, _ = r.Seek(0, 1)
	if err := binary.Read(r, binary.LittleEndian, &fib.CbRgFcLcb); err != nil {
		return nil, fmt.Errorf("fib: failed to read CbRgFcLcb at offset %d: %w", currentOffset, err)
	}

	// Read the variable-length FibRgFcLcb
	// CbRgFcLcb is a count of 64-bit values (8 bytes each).
	blobSize := int(fib.CbRgFcLcb) * 8
	if r.Len() < blobSize {
		return nil, fmt.Errorf("fib: data too short for RgFcLcbBlob, expected %d bytes, have %d", blobSize, r.Len())
	}
	fib.RgFcLcbBlob = make([]byte, blobSize)
	if _, err := r.Read(fib.RgFcLcbBlob); err != nil {
		return nil, fmt.Errorf("fib: could not read RgFcLcbBlob: %w", err)
	}

	// Parse the FibRgFcLcb structure based on nFib version
	if err := parseFibRgFcLcb(fib); err != nil {
		return nil, fmt.Errorf("fib: failed to parse FibRgFcLcb: %w", err)
	}

	return fib, nil
}

// parseFibRgFcLcb parses the variable FibRgFcLcb section based on the nFib version.
func parseFibRgFcLcb(fib *FileInformationBlock) error {
	if len(fib.RgFcLcbBlob) == 0 {
		return nil
	}

	r := bytes.NewReader(fib.RgFcLcbBlob)

	// For Word 97 format (nFib 0x00C1), parse the full FibRgFcLcb97 structure
	// The structure size should be 372 bytes (93 pairs * 4 bytes each = 744 bytes total, but we have 93*8=744 bytes from spec)
	// However, since binary.Read may not work reliably due to padding, parse manually
	switch fib.Base.NFib {
	case 0x00C1: // Word 97
		return parseFibRgFcLcb97(fib, r)
	case 0x00D9, 0x0101, 0x010C, 0x0112: // Word 2000, 2002, 2003, Enhanced 2003
		// For newer versions, we still parse as much as we can using the base structure
		return parseFibRgFcLcb97(fib, r)
	default:
		// For unknown versions, try to parse basic fields manually
		return parseBasicFcLcb(fib)
	}
}

// parseFibRgFcLcb97 manually parses the FibRgFcLcb97 structure to avoid padding issues.
func parseFibRgFcLcb97(fib *FileInformationBlock, r *bytes.Reader) error {
	// Each field is a pair of uint32 values (fc, lcb)
	// We'll parse them sequentially
	if r.Len() < 8 {
		return fmt.Errorf("not enough data for FibRgFcLcb97")
	}

	// Parse key fields manually
	fields := make([]uint32, r.Len()/4)
	for i := range fields {
		if err := binary.Read(r, binary.LittleEndian, &fields[i]); err != nil {
			return err
		}
	}

	// Map important fields by their index in the structure
	if len(fields) >= 2 {
		fib.RgFcLcb.FcStshfOrig = fields[0]
		fib.RgFcLcb.LcbStshfOrig = fields[1]
	}
	if len(fields) >= 4 {
		fib.RgFcLcb.FcStshf = fields[2]
		fib.RgFcLcb.LcbStshf = fields[3]
	}

	// Find FcClx and LcbClx (they should be at indices 66 and 67 for Word 97)
	// FcClx/LcbClx are at byte offset 264 = field index 66 (264/4 = 66)
	if len(fields) >= 68 {
		fib.RgFcLcb.FcClx = fields[66]
		fib.RgFcLcb.LcbClx = fields[67]
	}

	// Parse other important fields
	if len(fields) >= 20 {
		fib.RgFcLcb.FcPlcfsed = fields[12]
		fib.RgFcLcb.LcbPlcfsed = fields[13]
	}
	if len(fields) >= 24 {
		fib.RgFcLcb.FcPlcfhdd = fields[16]
		fib.RgFcLcb.LcbPlcfhdd = fields[17]
	}

	return nil
}

// parseBasicFcLcb attempts to parse basic FcClx/LcbClx fields for unknown FIB versions.
func parseBasicFcLcb(fib *FileInformationBlock) error {
	// For most Word versions, FcClx and LcbClx are at predictable offsets
	// This is a fallback for unknown nFib versions
	const fcClxOffset = 264 // Common offset for FcClx in FibRgFcLcb

	if len(fib.RgFcLcbBlob) >= fcClxOffset+8 {
		fib.RgFcLcb.FcClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset:])
		fib.RgFcLcb.LcbClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset+4:])
	}

	return nil
}

// IsEncrypted returns true if the document is encrypted.
func (fib *FileInformationBlock) IsEncrypted() bool {
	return (fib.Base.Flags1 & 0x0100) != 0 // fEncrypted flag
}

// IsObfuscated returns true if the document uses XOR obfuscation.
func (fib *FileInformationBlock) IsObfuscated() bool {
	return (fib.Base.Flags1 & 0x8000) != 0 // fObfuscated flag
}

// GetTableStreamName returns the name of the table stream to use.
func (fib *FileInformationBlock) GetTableStreamName() string {
	if (fib.Base.Flags1 & 0x0200) != 0 { // fWhichTblStm flag
		return "1Table"
	}
	return "0Table"
}
