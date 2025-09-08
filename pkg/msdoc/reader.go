package msdoc

import (
	"bytes"
	"encoding/binary"
	"time"
	"unicode/utf16"
)

// Text extracts the plain text content from the document.
func (d *Document) Text() (string, error) {
	// Determine which table stream to use based on the fWhichTblStm flag in the FIB.
	// Bit 9 of flags1 (FibBase offset 10).
	tableStreamName := "0Table"
	if d.fib.Base.Flags1&0x0200 != 0 {
		tableStreamName = "1Table"
	}

	tableStream, err := d.reader.ReadStream(tableStreamName)
	if err != nil {
		return "", err
	}

	// The piece table (PlcPcd) describes where text pieces are located.
	// Its location is defined in the FIB.
	clxOffset := d.fib.RgFcLcb.FcClx
	clxSize := d.fib.RgFcLcb.LcbClx

	// The Clx is a structure that starts with a PlcPcd.
	// We need to parse it to find the text pieces.
	if clxSize == 0 {
		return "", nil // No text content
	}
	clx := tableStream[clxOffset : clxOffset+clxSize]

	// The first byte of the Clx indicates it's a PlcPcd
	if clx[0] != 0x02 {
		return "", nil // Not a piece table
	}

	// Parse the PLC structure for piece descriptors (PCDs).
	// A PLC contains an array of CPs followed by an array of data.
	// For a PlcPcd, the data elements are PCDs.
	plcSize := len(clx) - 1
	pcdPlc := clx[1:]

	// According to MS-DOC spec, the size of the aCP array is calculated from total size.
	// Each PCD is 8 bytes. Each CP is 4 bytes.
	// n = (cbPlc - 4) / (cbData + 4) -> (plcSize - 4) / (8 + 4)
	numPcds := (plcSize - 4) / 12

	wordStream, err := d.reader.ReadStream("WordDocument")
	if err != nil {
		return "", err
	}

	var textBuilder bytes.Buffer

	for i := 0; i < numPcds; i++ {
		// Get the offset of the PCD in the PLC structure
		pcdOffset := 4*(numPcds+1) + (i * 8)
		pcd := pcdPlc[pcdOffset : pcdOffset+8]

		// The file character position (fc) is in the PCD
		fileCharPos := int32(pcd[2]) | int32(pcd[3])<<8 | int32(pcd[4])<<16 | int32(pcd[5])<<24

		// Determine if text is ANSI (fc is even) or Unicode (fc is odd and shifted)
		isUnicode := (fileCharPos & 0x40000000) != 0
		actualPos := fileCharPos & 0x3FFFFFFF

		// Get the start and end character positions (CPs) for this piece
		cpStart := binary.LittleEndian.Uint32(pcdPlc[i*4:])
		cpEnd := binary.LittleEndian.Uint32(pcdPlc[(i+1)*4:])
		charCount := cpEnd - cpStart

		if isUnicode {
			// Unicode text is stored at position / 2
			bytePos := uint32(actualPos / 2)
			byteCount := charCount * 2
			utf16bytes := wordStream[bytePos : bytePos+byteCount]

			// Convert UTF-16LE to UTF-8
			u16s := make([]uint16, charCount)
			for j := 0; j < int(charCount); j++ {
				u16s[j] = binary.LittleEndian.Uint16(utf16bytes[j*2:])
			}
			runes := utf16.Decode(u16s)
			textBuilder.WriteString(string(runes))
		} else {
			// ANSI text is stored at the given position
			bytePos := uint32(actualPos)
			byteCount := charCount
			ansiBytes := wordStream[bytePos : bytePos+byteCount]
			// This assumes CP-1252 encoding, a simple cast to string works for many chars.
			// A full solution would use a proper character encoding library.
			textBuilder.Write(ansiBytes)
		}
	}

	return textBuilder.String(), nil
}

// Metadata extracts high-level metadata from the document.
// In a complete implementation, this would parse the OLE SummaryInformation stream.
func (d *Document) Metadata() Metadata {
	// This is a stub implementation. A real one would parse the dedicated metadata stream.
	return Metadata{
		Title:   "N/A",
		Author:  "N/A",
		Created: time.Time{},
	}
}
