package msdoc

import (
    "fmt"
    "os"
    "time"

    "github.com/TalentFormula/msdoc/fib"
    "github.com/TalentFormula/msdoc/ole2"
)

// Document represents a loaded Microsoft Word .doc file.
type Document struct {
    file   *os.File
    reader *ole2.Reader
    fib    *fib.FileInformationBlock
}

// Metadata holds summary information about the document.
// Note: A full implementation requires parsing the "SummaryInformation" stream.
// This is a placeholder structure for the planned API.
type Metadata struct {
    Title   string
    Author  string
    Created time.Time
}

// Open reads and parses the given .doc file.
// It prepares the document for further operations like text extraction.
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

// Close closes the underlying .doc file.
func (d *Document) Close() error {
    if d.file != nil {
        return d.file.Close()
    }
    return nil
}
