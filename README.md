# MsDoc

> A Go implementation of the Microsoft Word .doc file format reader

## Overview

MsDoc is a Go library that implements the Microsoft Word Binary File Format (.doc) specification (MS-DOC). It allows you to extract text content and metadata from Word 97-2003 documents.

## Features

- ✅ **Text Extraction**: Extract plain text content from .doc files
- ✅ **Formatted Text**: Extract text with complete formatting information (fonts, colors, styles)
- ✅ **Metadata Reading**: Access comprehensive document properties (title, author, creation date, and more)
- ✅ **Format Support**: Word 97-2003 (.doc) files
- ✅ **OLE2 Parsing**: Full OLE2 compound document support
- ✅ **Unicode Support**: Handles both ANSI and Unicode text content
- ✅ **Piece Table Processing**: Correctly reconstructs fragmented text
- ✅ **Encryption Support**: Full support for encrypted and password-protected documents
- ✅ **Embedded Objects**: Extract images, charts, OLE objects, and other embedded content
- ✅ **VBA Macro Support**: Extract and decompile VBA macros and projects
- ✅ **Complete Metadata**: Full SummaryInformation and DocumentSummaryInformation parsing
- ✅ **Document Creation**: Create new .doc files from scratch
- ✅ **Document Modification**: Modify existing documents (text, formatting, metadata)
- ✅ **Write Support**: Full document creation and modification capabilities

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

	// Extract comprehensive metadata
	meta := doc.Metadata()
	fmt.Println("=== Metadata ===")
	fmt.Printf("Title: %s\n", meta.Title)
	fmt.Printf("Author: %s\n", meta.Author)
	fmt.Printf("Company: %s\n", meta.Company)
	fmt.Printf("Created: %s\n", meta.Created)
	fmt.Printf("Word Count: %d\n", meta.WordCount)
	fmt.Printf("Page Count: %d\n", meta.PageCount)

	// Check for encrypted documents
	if doc.IsEncrypted() {
		fmt.Println("Document is encrypted")
	}

	// Check for VBA macros
	if doc.HasMacros() {
		fmt.Println("Document contains VBA macros")
		modules, err := doc.GetAllVBAModules()
		if err == nil {
			fmt.Printf("VBA Modules: %v\n", modules)
		}
	}

	// Check for embedded objects
	if doc.HasEmbeddedObjects() {
		fmt.Println("Document contains embedded objects")
		objects, err := doc.GetEmbeddedObjects()
		if err == nil {
			for pos, obj := range objects {
				fmt.Printf("Object at %d: %s\n", pos, obj.GetObjectInfo())
			}
		}
	}

	// Extract formatted text runs
	textRuns, err := doc.GetFormattedText()
	if err == nil {
		fmt.Println("=== Formatted Text ===")
		for i, run := range textRuns {
			fmt.Printf("Run %d: %s\n", i, run.Text[:min(50, len(run.Text))])
			if run.CharProps != nil {
				fmt.Printf("  Bold: %t, Italic: %t, Font Size: %d\n", 
					run.CharProps.Bold, run.CharProps.Italic, run.CharProps.FontSize)
			}
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
```

## Encrypted Documents

```go
package main

import (
	"fmt"
	"log"

	"github.com/TalentFormula/msdoc/pkg/msdoc"
)

func main() {
	// Open an encrypted document with password
	doc, err := msdoc.OpenWithPassword("encrypted.doc", "password123")
	if err != nil {
		log.Fatalf("failed to open encrypted DOC: %v", err)
	}
	defer doc.Close()

	// Extract text (automatically decrypted)
	text, err := doc.Text()
	if err != nil {
		log.Fatalf("failed to extract text: %v", err)
	}
	fmt.Println("Decrypted text:", text)
}
```

## Creating Documents

```go
package main

import (
	"log"

	"github.com/TalentFormula/msdoc/pkg/msdoc"
	"github.com/TalentFormula/msdoc/formatting"
)

func main() {
	// Create a new document
	writer := msdoc.NewDocumentWriter()
	
	// Set document metadata
	writer.SetTitle("My Document")
	writer.SetAuthor("John Doe")
	writer.SetCompany("Acme Corp")
	writer.SetKeywords("example, document, msdoc")

	// Add plain text
	writer.AddParagraph("Hello, World!")

	// Add formatted text
	boldProps := &formatting.CharacterProperties{
		Bold:     true,
		FontSize: 24, // 12pt
		Color:    formatting.Color{Red: 255, Green: 0, Blue: 0},
	}
	writer.AddFormattedText("This is bold red text", boldProps, nil)

	// Add paragraph with custom formatting
	paraProps := &formatting.ParagraphProperties{
		Alignment:   formatting.AlignCenter,
		SpaceBefore: 240, // 12pt space before
		SpaceAfter:  240, // 12pt space after
	}
	writer.AddFormattedParagraph("Centered paragraph", nil, paraProps)

	// Insert page break
	writer.InsertPageBreak()
	writer.AddParagraph("This is on page 2")

	// Save the document
	err := writer.Save("output.doc")
	if err != nil {
		log.Fatalf("failed to save document: %v", err)
	}
	
	fmt.Println("Document created successfully!")
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
│   ├── doc.go              # Document interface (Open, OpenWithPassword)
│   ├── reader.go           # Text and metadata extraction
│   └── writer.go           # Document creation and modification
├── crypto/                 # Encryption and decryption support
│   ├── rc4.go              # RC4 cipher implementation
│   └── encryption.go       # Encryption header parsing
├── ole2/                   # OLE2 compound file parser
│   ├── reader.go           # Stream and directory reading
│   ├── writer.go           # Stream and directory writing
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
├── formatting/             # Advanced formatting support
│   └── formatting.go       # Character, paragraph, and section formatting
├── objects/                # Embedded object extraction
│   └── objects.go          # OLE objects, images, charts
├── macros/                 # VBA macro support
│   └── macros.go           # VBA project and module extraction
├── metadata/               # Complete metadata support
│   └── metadata.go         # SummaryInformation and DocumentSummaryInformation
├── writer/                 # Document creation and modification
│   └── writer.go           # Document writer implementation
└── cmd/msdocdump/          # CLI tool
```

## Implementation Status

### Core Components (✅ Complete)

- **OLE2 Reader**: Full compound file parsing with proper stream extraction
- **OLE2 Writer**: Complete compound document creation and modification
- **FIB Parser**: Complete File Information Block parsing for all Word versions
- **Piece Table**: Correct text reconstruction from piece descriptors
- **Text Extraction**: Unicode and ANSI text support with proper encoding
- **Stream Processing**: WordDocument, Table, and Data stream processors
- **Encryption Support**: Full RC4 decryption for password-protected documents

### Data Structures (✅ Complete)

- **Character Positions (CP)**: Zero-based character indexing
- **PLC Structures**: Plex parsing with validation
- **Piece Descriptors (PCD)**: Text fragment location and encoding
- **FKP Structures**: Formatted disk page parsing (CHPX/PAPX)
- **Section Properties (SEP)**: Section-level formatting information

### Advanced Features (✅ Complete)

- **Metadata Extraction**: Complete SummaryInformation and DocumentSummaryInformation parsing
- **Encryption Support**: Full decryption and password validation
- **Formatting Information**: Complete character, paragraph, and section formatting
- **Embedded Objects**: Full OLE object, image, and chart extraction
- **VBA Macro Support**: Complete VBA project and module extraction with decompilation
- **Document Creation**: Full document creation and modification capabilities
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

✅ **All Previous Limitations Have Been Resolved** ✅

This implementation now provides **complete** .doc file format support including:

- ✅ **Encrypted Documents**: Full support with RC4 decryption
- ✅ **Password Protection**: Complete password-based document access
- ✅ **Embedded Objects**: Full OLE object extraction (images, charts, files)
- ✅ **Complex Formatting**: Rich formatting preserved and exposed in API
- ✅ **VBA Macros**: Complete macro extraction and decompilation
- ✅ **Write Operations**: Full document creation and modification support
- ✅ **Complete Metadata**: All standard and custom properties supported

## API Reference

### Types

```go
type Document struct { /* ... */ }
type DocumentWriter struct { /* ... */ }

type Metadata struct {
    // Core properties
    Title               string
    Subject             string  
    Author              string
    Keywords            string
    Comments            string
    Template            string
    LastAuthor          string
    ApplicationName     string
    Created             time.Time
    LastSaved           time.Time
    LastPrinted         time.Time
    TotalEditTime       int64
    PageCount           int32
    WordCount           int32
    CharCount           int32
    CharCountWithSpaces int32
    
    // Extended properties
    Company             string
    Manager             string
    Category            string
    Language            int32
    CustomProperties    map[string]interface{}
    
    // And many more standard properties...
}

type TextRun struct {
    Text      string
    StartPos  uint32
    EndPos    uint32
    CharProps *CharacterProperties
    ParaProps *ParagraphProperties
}

type EmbeddedObject struct {
    Type      ObjectType
    Name      string
    ClassName string
    Data      []byte
    Size      int64
    Position  uint32
    IsLinked  bool
    LinkPath  string
}

type VBAProject struct {
    Name        string
    Description string
    Modules     map[string]*Module
    References  []*Reference
    Protected   bool
}

type CharacterProperties struct {
    FontName      string
    FontSize      uint16
    Bold          bool
    Italic        bool
    Underline     UnderlineType
    Strikethrough bool
    Color         Color
    // And many more formatting properties...
}
```

### Functions

```go
// Reading documents
func Open(filename string) (*Document, error)
func OpenWithPassword(filename, password string) (*Document, error)

// Document information
func (d *Document) Close() error
func (d *Document) IsEncrypted() bool
func (d *Document) HasMacros() bool
func (d *Document) HasEmbeddedObjects() bool

// Text extraction
func (d *Document) Text() (string, error)
func (d *Document) GetFormattedText() ([]*TextRun, error)

// Metadata extraction
func (d *Document) Metadata() *Metadata

// Embedded objects
func (d *Document) GetEmbeddedObjects() (map[uint32]*EmbeddedObject, error)
func (d *Document) GetEmbeddedObject(position uint32) (*EmbeddedObject, error)

// VBA macros
func (d *Document) GetVBAProject() (*VBAProject, error)
func (d *Document) GetVBACode(moduleName string) (string, error)
func (d *Document) GetAllVBAModules() ([]string, error)

// Document creation
func NewDocumentWriter() *DocumentWriter
func (dw *DocumentWriter) SetTitle(title string)
func (dw *DocumentWriter) SetAuthor(author string)
func (dw *DocumentWriter) AddText(text string)
func (dw *DocumentWriter) AddParagraph(text string)
func (dw *DocumentWriter) AddFormattedText(text string, charProps *CharacterProperties, paraProps *ParagraphProperties)
func (dw *DocumentWriter) Save(filename string) error
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

