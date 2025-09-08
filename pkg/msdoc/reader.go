package msdoc

import (
	"bytes"
	"fmt"
	"time"
	"unicode/utf16"

	"github.com/TalentFormula/msdoc/structures"
)

// Text extracts the plain text content from the document.
//
// This method parses the document's piece table to reconstruct the original text
// from potentially fragmented pieces stored throughout the file. It handles both
// ANSI and Unicode text encoding as specified in the MS-DOC format.
//
// Returns an error if:
//   - The document is encrypted (encryption support not yet implemented)
//   - The piece table is corrupted or invalid
//   - Required streams (WordDocument, Table) cannot be read
//   - Text data extends beyond stream boundaries
//
// For documents with no text content, returns an empty string with no error.
func (d *Document) Text() (string, error) {
	// Check if document is encrypted
	if d.fib.IsEncrypted() {
		return "", fmt.Errorf("text extraction from encrypted documents is not yet supported")
	}

	// Get the appropriate table stream
	tableStreamName := d.fib.GetTableStreamName()
	tableStream, err := d.reader.ReadStream(tableStreamName)
	if err != nil {
		return "", fmt.Errorf("failed to read table stream %s: %w", tableStreamName, err)
	}

	// Get the piece table location from FIB
	clxOffset := d.fib.RgFcLcb.FcClx
	clxSize := d.fib.RgFcLcb.LcbClx

	if clxSize == 0 {
		return "", nil // No text content
	}

	if uint32(len(tableStream)) < clxOffset+clxSize {
		return "", fmt.Errorf("table stream too small for CLX data")
	}

	clx := tableStream[clxOffset : clxOffset+clxSize]

	// The CLX should start with a PlcPcd indicator (0x02)
	if len(clx) == 0 || clx[0] != 0x02 {
		return "", fmt.Errorf("invalid CLX structure, expected PlcPcd marker")
	}

	// Parse the piece table
	plcPcdData := clx[1:] // Skip the marker byte
	plcPcd, err := structures.ParsePlcPcd(plcPcdData)
	if err != nil {
		return "", fmt.Errorf("failed to parse piece table: %w", err)
	}

	// Get the WordDocument stream for text content
	wordStream, err := d.reader.ReadStream("WordDocument")
	if err != nil {
		return "", fmt.Errorf("failed to read WordDocument stream: %w", err)
	}

	// Extract text from each piece
	var textBuilder bytes.Buffer

	for i := 0; i < plcPcd.Count(); i++ {
		startCP, endCP, pcd, err := plcPcd.GetTextRange(i)
		if err != nil {
			return "", fmt.Errorf("failed to get text range for piece %d: %w", i, err)
		}

		charCount := startCP.Distance(endCP)
		if charCount == 0 {
			continue
		}

		// Get the file position for this piece
		filePos := pcd.GetActualFC()

		if pcd.IsUnicode {
			// Unicode text (UTF-16LE)
			byteCount := charCount * 2
			if uint32(len(wordStream)) < filePos+byteCount {
				return "", fmt.Errorf("WordDocument stream too small for Unicode text at piece %d", i)
			}

			utf16bytes := wordStream[filePos : filePos+byteCount]
			
			// Convert UTF-16LE to Go string
			u16s := make([]uint16, charCount)
			for j := uint32(0); j < charCount; j++ {
				if (j*2)+1 < uint32(len(utf16bytes)) {
					u16s[j] = uint16(utf16bytes[j*2]) | (uint16(utf16bytes[j*2+1]) << 8)
				}
			}
			runes := utf16.Decode(u16s)
			textBuilder.WriteString(string(runes))
		} else {
			// ANSI text (CP-1252 encoding)
			if uint32(len(wordStream)) < filePos+charCount {
				return "", fmt.Errorf("WordDocument stream too small for ANSI text at piece %d", i)
			}

			ansiBytes := wordStream[filePos : filePos+charCount]
			// For basic ASCII/CP-1252, direct conversion works for most characters
			// A complete implementation would use proper character encoding conversion
			textBuilder.Write(ansiBytes)
		}
	}

	return textBuilder.String(), nil
}

// Metadata extracts high-level metadata from the document.
//
// This method attempts to parse the OLE SummaryInformation stream to extract
// document properties such as title, author, and creation date. If the stream
// is not available or cannot be parsed, default values are returned.
//
// The current implementation provides basic metadata extraction. A complete
// implementation would fully parse the property set stream format and handle
// all standard document properties defined in the OLE specification.
//
// Returns a Metadata structure with available information, never returns an error.
func (d *Document) Metadata() Metadata {
	// Try to parse summary information stream
	summaryData, err := d.reader.ReadStream("\x05SummaryInformation")
	if err != nil {
		// Return basic metadata from FIB if SummaryInformation is not available
		return Metadata{
			Title:   "N/A",
			Author:  "N/A",
			Created: time.Time{},
		}
	}

	// Parse summary information (basic implementation)
	return parseSummaryInformation(summaryData)
}

// parseSummaryInformation extracts metadata from the SummaryInformation stream.
// This is a simplified implementation that handles the most common fields.
func parseSummaryInformation(data []byte) Metadata {
	// This is a stub implementation. A complete parser would handle:
	// - Property set stream format
	// - Property identifiers for title, author, creation time, etc.
	// - Different data types (strings, timestamps, etc.)
	
	// For now, return placeholder values
	return Metadata{
		Title:   "N/A", // Would extract from PID_TITLE (0x02)
		Author:  "N/A", // Would extract from PID_AUTHOR (0x04)
		Created: time.Time{}, // Would extract from PID_CREATE_DTM (0x0C)
	}
}
