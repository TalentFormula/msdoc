// Package msdoc provides functionality for reading and parsing Microsoft Word .doc files.
//
// This package implements the MS-DOC binary file format specification, allowing
// extraction of text content and metadata from Word 97-2003 documents.
//
// Basic usage:
//
//	doc, err := msdoc.Open("document.doc")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer doc.Close()
//
//	text, err := doc.Text()
//	if err != nil {
//		log.Fatal(err)
//	}
//	fmt.Println(text)
//
//	metadata := doc.Metadata()
//	fmt.Printf("Title: %s\n", metadata.Title)
//	fmt.Printf("Author: %s\n", metadata.Author)
package msdoc

import (
	"fmt"
	"os"
	"time"

	"github.com/TalentFormula/msdoc/fib"
	"github.com/TalentFormula/msdoc/ole2"
)

// Document represents a loaded Microsoft Word .doc file.
// It provides methods for extracting text content and metadata.
type Document struct {
	file   *os.File
	reader *ole2.Reader
	fib    *fib.FileInformationBlock
}

// Metadata holds summary information about the document.
// The fields are extracted from the SummaryInformation stream when available.
type Metadata struct {
	Title   string    // Document title
	Author  string    // Document author
	Created time.Time // Creation timestamp
}

// Open reads and parses the given .doc file.
// It prepares the document for further operations like text extraction.
//
// The file must be a valid Microsoft Word .doc file (Word 97-2003 format).
// Encrypted documents are detected but not currently supported for text extraction.
//
// Returns an error if the file cannot be opened, is not a valid .doc file,
// or if the internal OLE2 structure is corrupted.
func Open(filename string) (*Document, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filename, err)
	}

	oleReader, err := ole2.NewReader(file)
	if err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to create OLE2 reader: %w", err)
	}

	// The FIB is located in the "WordDocument" stream.
	wordDocumentStream, err := oleReader.ReadStream("WordDocument")
	if err != nil {
		file.Close()
		return nil, fmt.Errorf("could not find WordDocument stream: %w", err)
	}

	fib, err := fib.ParseFIB(wordDocumentStream)
	if err != nil {
		file.Close()
		return nil, fmt.Errorf("failed to parse FIB: %w", err)
	}

	doc := &Document{
		file:   file,
		reader: oleReader,
		fib:    fib,
	}

	return doc, nil
}

// Close closes the underlying .doc file and releases associated resources.
// It is safe to call Close multiple times.
func (d *Document) Close() error {
	if d.file != nil {
		return d.file.Close()
	}
	return nil
}
