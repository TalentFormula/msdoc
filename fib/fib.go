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

	// Move the reader back to the start to parse the whole structure
	r.Seek(0, 0)

	if err := binary.Read(r, binary.LittleEndian, &fib.Base); err != nil {
		return nil, err
	}
	if err := binary.Read(r, binary.LittleEndian, &fib.Csw); err != nil {
		return nil, err
	}
	if err := binary.Read(r, binary.LittleEndian, &fib.FibRgW); err != nil {
		return nil, err
	}
	if err := binary.Read(r, binary.LittleEndian, &fib.Cslw); err != nil {
		return nil, err
	}
	if err := binary.Read(r, binary.LittleEndian, &fib.FibRgLw); err != nil {
		return nil, err
	}
	if err := binary.Read(r, binary.LittleEndian, &fib.CbRgFcLcb); err != nil {
		return nil, err
	}

	// Read the variable-length FibRgFcLcb
	// CbRgFcLcb is a count of 64-bit values (8 bytes).
	blobSize := int(fib.CbRgFcLcb) * 8
	if r.Len() < blobSize {
		return nil, fmt.Errorf("fib: data too short for RgFcLcbBlob, expected %d bytes, have %d", blobSize, r.Len())
	}
	fib.RgFcLcbBlob = make([]byte, blobSize)
	if _, err := r.Read(fib.RgFcLcbBlob); err != nil {
		return nil, fmt.Errorf("fib: could not read RgFcLcbBlob: %w", err)
	}

	// Parse known fields from RgFcLcbBlob for text extraction.
	// For nFib=0x00C1 (Word 97), the FibRgFcLcb97 structure is used.
	// The fcClx field is at offset 0x108 (264) within that structure.
	const fcClxOffset = 264
	if len(fib.RgFcLcbBlob) >= fcClxOffset+8 {
		fib.RgFcLcb.FcClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset:])
		fib.RgFcLcb.LcbClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset+4:])
	}

	return fib, nil
}
