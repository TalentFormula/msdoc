// Package msdoc provides comprehensive support for creating and modifying .doc files.
//
// This package now supports full write operations including document creation,
// text insertion, formatting application, and complete OLE2 compound document
// generation according to the MS-DOC specification.
//
// Example usage:
//
//	writer := msdoc.NewWriter()
//	writer.SetTitle("My Document")
//	writer.SetAuthor("John Doe")
//	writer.AddParagraph("Hello, World!")
//
//	// Add formatted text
//	charProps := &formatting.CharacterProperties{
//		Bold: true,
//		FontSize: 24, // 12pt
//	}
//	writer.AddFormattedText("Bold text", charProps, nil)
//
//	err := writer.Save("output.doc")
//
// This implementation provides complete document creation capabilities including
// text content, formatting, metadata, and proper OLE2 structure generation.
package msdoc

import (
	"github.com/TalentFormula/msdoc/writer"
)

// DocumentWriter provides functionality for creating and modifying .doc files.
// This is an alias for writer.DocumentWriter to maintain clean public API.
type DocumentWriter = writer.DocumentWriter

// NewDocumentWriter creates a new document writer for creating .doc files.
// This function replaces the previous stub implementation with full functionality.
func NewDocumentWriter() *DocumentWriter {
	return writer.NewDocumentWriter()
}
