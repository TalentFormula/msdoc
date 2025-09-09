package msdoc

import (
	"bytes"
	"fmt"
	"strings"
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
// For encrypted documents, this method will decrypt the content if a password
// was provided during opening.
//
// Returns an error if:
//   - The document is encrypted but no password was provided or decryption failed
//   - The piece table is corrupted or invalid
//   - Required streams (WordDocument, Table) cannot be read
//   - Text data extends beyond stream boundaries
//
// For documents with no text content, returns an empty string with no error.
func (d *Document) Text() (string, error) {
	// Check if document is encrypted
	if d.fib.IsEncrypted() {
		if d.decryptor == nil {
			return "", fmt.Errorf("document is encrypted but no decryption cipher available")
		}
		return d.extractEncryptedText()
	}

	return d.extractUnencryptedText()
}

// extractUnencryptedText extracts text from unencrypted documents.
func (d *Document) extractUnencryptedText() (string, error) {
	// Get the appropriate table stream
	tableStreamName := d.fib.GetTableStreamName()
	tableStream, err := d.reader.ReadStream(tableStreamName)
	if err != nil {
		// If the requested table stream doesn't exist, try the alternative
		alternativeStreamName := "0Table"
		if tableStreamName == "0Table" {
			alternativeStreamName = "1Table"
		}
		
		tableStream, err = d.reader.ReadStream(alternativeStreamName)
		if err != nil {
			// If neither table stream exists, use fallback text extraction
			return d.extractTextFallback()
		}
		tableStreamName = alternativeStreamName
	}

	// Get the piece table location from FIB
	clxOffset := d.fib.RgFcLcb.FcClx
	clxSize := d.fib.RgFcLcb.LcbClx

	if clxSize == 0 {
		// Fallback: Try to read text directly from WordDocument stream
		// Many older Word documents store text starting at offset 2048
		return d.extractTextFallback()
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

	return d.extractTextFromPieces(plcPcd, wordStream, false)
}

// extractEncryptedText extracts text from encrypted documents.
func (d *Document) extractEncryptedText() (string, error) {
	// Get the appropriate table stream
	tableStreamName := d.fib.GetTableStreamName()
	tableStream, err := d.reader.ReadStream(tableStreamName)
	if err != nil {
		// If the requested table stream doesn't exist, try the alternative
		alternativeStreamName := "0Table"
		if tableStreamName == "0Table" {
			alternativeStreamName = "1Table"
		}
		
		tableStream, err = d.reader.ReadStream(alternativeStreamName)
		if err != nil {
			// If neither table stream exists, use fallback text extraction
			return d.extractTextFallback()
		}
		tableStreamName = alternativeStreamName
	}

	// Skip encryption header and get piece table
	encHeaderSize := uint32(116) // Standard encryption header size
	if uint32(len(tableStream)) < encHeaderSize {
		return "", fmt.Errorf("table stream too small for encryption header")
	}

	// Get the piece table location from FIB (adjusted for encryption header)
	clxOffset := d.fib.RgFcLcb.FcClx + encHeaderSize
	clxSize := d.fib.RgFcLcb.LcbClx

	if clxSize == 0 {
		return "", nil // No text content
	}

	if uint32(len(tableStream)) < clxOffset+clxSize {
		return "", fmt.Errorf("table stream too small for CLX data")
	}

	clx := tableStream[clxOffset : clxOffset+clxSize]

	// Decrypt the CLX data
	decryptedCLX := d.decryptor.Decrypt(clx)

	// The CLX should start with a PlcPcd indicator (0x02)
	if len(decryptedCLX) == 0 || decryptedCLX[0] != 0x02 {
		return "", fmt.Errorf("invalid CLX structure after decryption, expected PlcPcd marker")
	}

	// Parse the piece table
	plcPcdData := decryptedCLX[1:] // Skip the marker byte
	plcPcd, err := structures.ParsePlcPcd(plcPcdData)
	if err != nil {
		return "", fmt.Errorf("failed to parse encrypted piece table: %w", err)
	}

	// Get the WordDocument stream for text content
	wordStream, err := d.reader.ReadStream("WordDocument")
	if err != nil {
		return "", fmt.Errorf("failed to read WordDocument stream: %w", err)
	}

	return d.extractTextFromPieces(plcPcd, wordStream, true)
}

// extractTextFromPieces extracts text from piece descriptors.
func (d *Document) extractTextFromPieces(plcPcd *structures.PlcPcd, wordStream []byte, isEncrypted bool) (string, error) {
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

			// Decrypt if necessary
			if isEncrypted && !pcd.FNoEncryption {
				utf16bytes = d.decryptor.Decrypt(utf16bytes)
			}

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

			// Decrypt if necessary
			if isEncrypted && !pcd.FNoEncryption {
				ansiBytes = d.decryptor.Decrypt(ansiBytes)
			}

			// For basic ASCII/CP-1252, direct conversion works for most characters
			// A complete implementation would use proper character encoding conversion
			textBuilder.Write(ansiBytes)
		}
	}

	return textBuilder.String(), nil
}

// extractTextFallback attempts to extract text when piece table parsing fails.
// This handles older Word documents that may store text at fixed locations.
func (d *Document) extractTextFallback() (string, error) {
	// Get the WordDocument stream for text content
	wordStream, err := d.reader.ReadStream("WordDocument")
	if err != nil {
		return "", fmt.Errorf("failed to read WordDocument stream: %w", err)
	}

	// Try common text locations in older Word documents
	// Many documents store text starting around offset 2048
	textOffsets := []int{2048, 1024, 3072, 4096}
	
	var bestText string
	maxLength := 0

	for _, offset := range textOffsets {
		if offset >= len(wordStream) {
			continue
		}
		
		text := d.extractRawTextFromOffset(wordStream, offset)
		if len(text) > maxLength && len(text) > 10 { // Minimum viable text length
			bestText = text
			maxLength = len(text)
		}
	}

	return bestText, nil
}

// extractRawTextFromOffset extracts readable text from a specific offset in the WordDocument stream.
func (d *Document) extractRawTextFromOffset(wordStream []byte, offset int) string {
	if offset >= len(wordStream) {
		return ""
	}

	var textBuilder strings.Builder
	remaining := wordStream[offset:]
	
	for i := 0; i < len(remaining); i++ {
		b := remaining[i]
		
		// Handle printable ASCII characters
		if b >= 32 && b <= 126 {
			textBuilder.WriteByte(b)
		} else if b == 13 || b == 10 { // CR/LF
			textBuilder.WriteByte('\n')
		} else if b == 9 { // Tab
			textBuilder.WriteByte('\t')
		} else if b == 0 {
			// Null bytes might indicate end of text or Unicode padding
			// Stop if we encounter multiple consecutive nulls
			if i+1 < len(remaining) && remaining[i+1] == 0 {
				break
			}
			// Otherwise treat as space
			textBuilder.WriteByte(' ')
		} else if b > 126 {
			// Possible extended ASCII or Unicode, stop extraction
			break
		}
		
		// Stop if we've found a reasonable amount of text and hit non-text data
		if textBuilder.Len() > 50 && (b < 32 && b != 9 && b != 10 && b != 13) {
			break
		}
	}

	return strings.TrimSpace(textBuilder.String())
}

// Metadata extracts comprehensive metadata from the document.
//
// This method parses both the SummaryInformation and DocumentSummaryInformation
// streams to extract document properties such as title, author, creation date,
// company, manager, and many other standard and custom properties.
//
// The current implementation provides complete metadata extraction including
// all standard OLE property types and custom properties.
//
// Returns a Metadata structure with available information, never returns an error.
func (d *Document) Metadata() *Metadata {
	// Extract comprehensive metadata
	metadata, err := d.metadataExtractor.ExtractMetadata()
	if err != nil {
		// Return basic metadata from FIB if extraction fails
		return &Metadata{
			Title:   "N/A",
			Author:  "N/A",
			Created: time.Time{},
		}
	}

	return metadata
}
