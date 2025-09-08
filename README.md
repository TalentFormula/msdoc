# MsDoc

> A Go implementation of the Microsoft Word .doc file format reader

## Overview

MsDoc is a Go library that implements the Microsoft Word Binary File Format (.doc) specification (MS-DOC). It allows you to extract text content and metadata from Word 97-2003 documents.

## Features

- ✅ **Text Extraction**: Extract plain text content from .doc files
- ✅ **Metadata Reading**: Access document properties (title, author, creation date)
- ✅ **Format Support**: Word 97-2003 (.doc) files
- ✅ **OLE2 Parsing**: Full OLE2 compound document support
- ✅ **Unicode Support**: Handles both ANSI and Unicode text content
- ✅ **Piece Table Processing**: Correctly reconstructs fragmented text
- ⚠️ **Encryption Detection**: Detects encrypted files (extraction not yet supported)
- ❌ **Formatting Preservation**: Formatting information is parsed but not exposed in public API
- ❌ **Write Support**: Read-only library (no document creation/modification)

## Installation

```bash
go get github.com/TalentFormula/msdoc
```

## Quick Start

```go
package main

import (
	"fmt"
	"log"

	"github.com/TalentFormula/msdoc/pkg/msdoc"
)

func main() {
	// Open a .doc file
	doc, err := msdoc.Open("sample.doc")
	if err != nil {
		log.Fatalf("failed to open DOC: %v", err)
	}
	defer doc.Close()

	// Extract plain text
	text, err := doc.Text()
	if err != nil {
		log.Fatalf("failed to extract text: %v", err)
	}
	fmt.Println("=== Document Text ===")
	fmt.Println(text)

	// Extract metadata
	meta := doc.Metadata()
	fmt.Println("=== Metadata ===")
	fmt.Printf("Title: %s\n", meta.Title)
	fmt.Printf("Author: %s\n", meta.Author)
	fmt.Printf("Created: %s\n", meta.Created)
}
```

## Command Line Tool

The library includes a command-line tool for debugging and text extraction:

```bash
go install github.com/TalentFormula/msdoc/cmd/msdocdump
msdocdump document.doc
```

## Architecture

The library is structured according to the MS-DOC specification:

```
msdoc/
├── pkg/msdoc/              # Public API
│   ├── doc.go              # Document interface (Open, Close)
│   ├── reader.go           # Text and metadata extraction
│   └── writer.go           # Future write support (stub)
├── ole2/                   # OLE2 compound file parser
│   ├── reader.go           # Stream and directory reading
│   ├── sector.go           # Sector management
│   └── directory.go        # Directory entry parsing
├── fib/                    # File Information Block parsing
│   ├── fib.go              # FIB parser
│   └── fib_types.go        # FIB structure definitions
├── streams/                # Stream processors
│   ├── worddocument.go     # WordDocument stream
│   ├── table.go            # Table stream (0Table/1Table)
│   └── data.go             # Data stream
├── structures/             # DOC data structures
│   ├── cp.go               # Character Position handling
│   ├── plc.go              # PLC (Plex) structures
│   ├── pcd.go              # Piece Descriptor structures
│   ├── fkp.go              # Formatted disk pages
│   └── sep.go              # Section properties
└── cmd/msdocdump/          # CLI tool
```

## Implementation Status

### Core Components (✅ Complete)

- **OLE2 Reader**: Full compound file parsing with proper stream extraction
- **FIB Parser**: Complete File Information Block parsing for all Word versions
- **Piece Table**: Correct text reconstruction from piece descriptors
- **Text Extraction**: Unicode and ANSI text support with proper encoding
- **Stream Processing**: WordDocument, Table, and Data stream processors

### Data Structures (✅ Complete)

- **Character Positions (CP)**: Zero-based character indexing
- **PLC Structures**: Plex parsing with validation
- **Piece Descriptors (PCD)**: Text fragment location and encoding
- **FKP Structures**: Formatted disk page parsing (CHPX/PAPX)
- **Section Properties (SEP)**: Section-level formatting information

### Advanced Features (⚠️ Partial)

- **Metadata Extraction**: Basic framework (needs SummaryInformation parser)
- **Encryption Support**: Detection only (decryption not implemented)
- **Formatting Information**: Parsed but not exposed in public API
- **Error Handling**: Comprehensive validation and error reporting

## Supported Document Versions

| Version | nFib Code | Support Status |
|---------|-----------|---------------|
| Word 97 | 0x00C1 | ✅ Full support |
| Word 2000 | 0x00D9 | ✅ Full support |
| Word 2002 (XP) | 0x0101 | ✅ Full support |
| Word 2003 | 0x010C | ✅ Full support |
| Word 2003 Enhanced | 0x0112 | ✅ Full support |

## Limitations

- **Encrypted Documents**: Detection only, decryption not supported
- **Password Protection**: Cannot access password-protected documents
- **Embedded Objects**: OLE objects are not extracted
- **Complex Formatting**: Rich formatting is not preserved in text output
- **Macros**: VBA macros are not processed or extracted
- **Write Operations**: Library is read-only

## API Reference

### Types

```go
type Document struct { /* ... */ }
type Metadata struct {
    Title   string
    Author  string
    Created time.Time
}
```

### Functions

```go
// Open a .doc file for reading
func Open(filename string) (*Document, error)

// Close the document and release resources
func (d *Document) Close() error

// Extract all text content as a single string
func (d *Document) Text() (string, error)

// Get document metadata
func (d *Document) Metadata() Metadata
```

## Testing

```bash
# Run all tests
go test ./...

# Run tests with verbose output
go test -v ./...

# Test specific components
go test ./ole2
go test ./fib
go test ./structures
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Ensure all tests pass
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE.txt file for details.

## References

- [MS-DOC Specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/): Official Microsoft specification
- [MS-CFB Specification](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/): OLE2 Compound File Binary Format
- [wvWare Project](http://wvware.sourceforge.net/): Historical reverse engineering efforts

## Acknowledgments

This implementation is based on the official Microsoft Open Specifications Program documentation and incorporates insights from various open-source projects that have worked with the .doc format over the years.

