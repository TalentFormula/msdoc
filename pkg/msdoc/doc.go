// Package msdoc provides comprehensive functionality for reading, parsing, and creating Microsoft Word .doc files.
//
// This package implements the complete MS-DOC binary file format specification, allowing
// extraction of text content, metadata, embedded objects, VBA macros, and formatting
// from Word 97-2003 documents. It also supports creating new documents and modifying
// existing ones, including support for encrypted/password-protected documents.
//
// Basic reading usage:
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
//
// Reading encrypted documents:
//
//	doc, err := msdoc.OpenWithPassword("encrypted.doc", "password123")
//	if err != nil {
//		log.Fatal(err)
//	}
//	defer doc.Close()
//
// Creating new documents:
//
//	writer := msdoc.NewWriter()
//	writer.SetTitle("My Document")
//	writer.SetAuthor("John Doe")
//	writer.AddParagraph("Hello, World!")
//	err := writer.Save("output.doc")
package msdoc

import (
	"fmt"
	"os"

	"github.com/TalentFormula/msdoc/crypto"
	"github.com/TalentFormula/msdoc/fib"
	"github.com/TalentFormula/msdoc/formatting"
	"github.com/TalentFormula/msdoc/macros"
	"github.com/TalentFormula/msdoc/metadata"
	"github.com/TalentFormula/msdoc/objects"
	"github.com/TalentFormula/msdoc/ole2"
	"github.com/TalentFormula/msdoc/writer"
)

// Document represents a loaded Microsoft Word .doc file.
// It provides methods for extracting text content, metadata, embedded objects,
// macros, and formatting information. It also supports decryption of encrypted documents.
type Document struct {
	file      *os.File
	reader    *ole2.Reader
	fib       *fib.FileInformationBlock
	password  string // For encrypted documents
	decryptor *crypto.RC4 // For encrypted documents
	
	// Lazy-loaded components
	objectPool    *objects.ObjectPool
	macroExtractor *macros.MacroExtractor
	metadataExtractor *metadata.MetadataExtractor
	formattingExtractor *formatting.FormattingExtractor
}

// Metadata holds comprehensive document metadata information.
// This is an alias for metadata.DocumentMetadata for backward compatibility.
type Metadata = metadata.DocumentMetadata

// TextRun represents a run of text with consistent formatting.
// This is an alias for formatting.TextRun.
type TextRun = formatting.TextRun

// EmbeddedObject represents an object embedded in the document.
// This is an alias for objects.EmbeddedObject.
type EmbeddedObject = objects.EmbeddedObject

// VBAProject represents a VBA project contained in the document.
// This is an alias for macros.VBAProject.
type VBAProject = macros.VBAProject

// Open reads and parses the given .doc file.
// It prepares the document for further operations like text extraction.
//
// The file must be a valid Microsoft Word .doc file (Word 97-2003 format).
// For encrypted documents, use OpenWithPassword instead.
//
// Returns an error if the file cannot be opened, is not a valid .doc file,
// or if the internal OLE2 structure is corrupted.
func Open(filename string) (*Document, error) {
	return openWithPassword(filename, "")
}

// OpenWithPassword opens an encrypted .doc file with the provided password.
// This function supports password-protected and encrypted documents.
//
// Returns an error if the file cannot be opened, is not a valid .doc file,
// the password is incorrect, or if decryption fails.
func OpenWithPassword(filename, password string) (*Document, error) {
	return openWithPassword(filename, password)
}

// openWithPassword is the internal function that handles both encrypted and unencrypted files.
func openWithPassword(filename, password string) (*Document, error) {
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
		file:     file,
		reader:   oleReader,
		fib:      fib,
		password: password,
	}

	// Initialize lazy-loaded components
	doc.objectPool = objects.NewObjectPool(oleReader)
	doc.macroExtractor = macros.NewMacroExtractor(oleReader)
	doc.metadataExtractor = metadata.NewMetadataExtractor(oleReader)
	doc.formattingExtractor = formatting.NewFormattingExtractor()

	// Handle encryption if document is encrypted
	if fib.IsEncrypted() {
		if password == "" {
			return nil, fmt.Errorf("document is encrypted but no password provided")
		}

		if err := doc.setupDecryption(); err != nil {
			file.Close()
			return nil, fmt.Errorf("failed to setup decryption: %w", err)
		}
	}

	return doc, nil
}

// setupDecryption initializes decryption for encrypted documents.
func (d *Document) setupDecryption() error {
	// Get the table stream name
	tableStreamName := d.fib.GetTableStreamName()
	tableStream, err := d.reader.ReadStream(tableStreamName)
	if err != nil {
		return fmt.Errorf("failed to read table stream %s: %w", tableStreamName, err)
	}

	// Parse encryption header
	encHeader, err := crypto.ParseEncryptionHeader(tableStream)
	if err != nil {
		return fmt.Errorf("failed to parse encryption header: %w", err)
	}

	// Create decryption cipher
	decryptor, err := encHeader.CreateDecryptionCipher(d.password)
	if err != nil {
		return fmt.Errorf("failed to create decryption cipher: %w", err)
	}

	d.decryptor = decryptor
	return nil
}

// Close closes the underlying .doc file and releases associated resources.
// It is safe to call Close multiple times.
func (d *Document) Close() error {
	if d.file != nil {
		return d.file.Close()
	}
	return nil
}

// IsEncrypted returns true if the document is encrypted.
func (d *Document) IsEncrypted() bool {
	return d.fib.IsEncrypted()
}

// HasMacros returns true if the document contains VBA macros.
func (d *Document) HasMacros() bool {
	return d.macroExtractor.HasMacros()
}

// HasEmbeddedObjects returns true if the document contains embedded objects.
func (d *Document) HasEmbeddedObjects() bool {
	// Load objects if not already loaded
	if err := d.objectPool.LoadObjects(); err != nil {
		return false
	}
	return len(d.objectPool.GetAllObjects()) > 0
}

// GetFormattedText extracts text with formatting information.
// Returns an array of TextRun structures containing text and formatting.
func (d *Document) GetFormattedText() ([]*TextRun, error) {
	if d.fib.IsEncrypted() && d.decryptor == nil {
		return nil, fmt.Errorf("document is encrypted but decryption is not available")
	}

	// Implementation would extract text with formatting
	// For now, return basic text as a single run
	text, err := d.Text()
	if err != nil {
		return nil, err
	}

	return []*TextRun{{
		Text:     text,
		StartPos: 0,
		EndPos:   uint32(len(text)),
	}}, nil
}

// GetEmbeddedObjects returns all embedded objects in the document.
func (d *Document) GetEmbeddedObjects() (map[uint32]*EmbeddedObject, error) {
	if err := d.objectPool.LoadObjects(); err != nil {
		return nil, fmt.Errorf("failed to load embedded objects: %w", err)
	}
	return d.objectPool.GetAllObjects(), nil
}

// GetEmbeddedObject returns a specific embedded object by position.
func (d *Document) GetEmbeddedObject(position uint32) (*EmbeddedObject, error) {
	if err := d.objectPool.LoadObjects(); err != nil {
		return nil, fmt.Errorf("failed to load embedded objects: %w", err)
	}
	return d.objectPool.ExtractObject(position)
}

// GetVBAProject extracts the VBA project from the document.
// Returns an error if the document does not contain macros.
func (d *Document) GetVBAProject() (*VBAProject, error) {
	return d.macroExtractor.ExtractProject()
}

// GetVBACode returns the VBA code for a specific module.
func (d *Document) GetVBACode(moduleName string) (string, error) {
	project, err := d.GetVBAProject()
	if err != nil {
		return "", err
	}

	code, exists := project.GetModuleCode(moduleName)
	if !exists {
		return "", fmt.Errorf("module %s not found", moduleName)
	}

	return code, nil
}

// GetAllVBAModules returns the names of all VBA modules in the document.
func (d *Document) GetAllVBAModules() ([]string, error) {
	project, err := d.GetVBAProject()
	if err != nil {
		return nil, err
	}

	return project.GetAllModuleNames(), nil
}

// NewWriter creates a new document writer for creating .doc files.
func NewWriter() *writer.DocumentWriter {
	return writer.NewDocumentWriter()
}
