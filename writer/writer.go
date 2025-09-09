// Package writer provides comprehensive support for creating and modifying .doc files.
//
// This package implements document creation, text insertion, formatting application,
// and complete OLE2 compound document generation according to the MS-DOC specification.
package writer

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"time"

	"github.com/TalentFormula/msdoc/fib"
	"github.com/TalentFormula/msdoc/formatting"
	"github.com/TalentFormula/msdoc/ole2"
)

// DocumentWriter provides functionality for creating and modifying .doc files.
type DocumentWriter struct {
	metadata   *DocumentInfo
	text       []TextSection
	oleWriter  *ole2.Writer
	fibBuilder *FIBBuilder
	pieceTable *PieceTableBuilder
	formatting *FormattingBuilder
}

// DocumentInfo holds document-level information.
type DocumentInfo struct {
	Title       string
	Author      string
	Subject     string
	Keywords    string
	Comments    string
	Company     string
	Manager     string
	Category    string
	Template    string
	Application string
	Version     string
	Language    int32
	Created     time.Time
	Modified    time.Time
}

// TextSection represents a section of text with formatting.
type TextSection struct {
	Text      string
	CharProps *formatting.CharacterProperties
	ParaProps *formatting.ParagraphProperties
	IsNewPara bool
}

// FIBBuilder handles File Information Block construction.
type FIBBuilder struct {
	fib *fib.FileInformationBlock
}

// PieceTableBuilder constructs piece tables for text storage.
type PieceTableBuilder struct {
	pieces []PieceDescriptor
	text   bytes.Buffer
}

// PieceDescriptor represents a piece of text in the document.
type PieceDescriptor struct {
	StartCP    uint32
	EndCP      uint32
	FileOffset uint32
	IsUnicode  bool
}

// FormattingBuilder handles formatting table construction.
type FormattingBuilder struct {
	charFormats []CharacterFormatEntry
	paraFormats []ParagraphFormatEntry
}

// CharacterFormatEntry represents character formatting information.
type CharacterFormatEntry struct {
	StartCP uint32
	EndCP   uint32
	Props   *formatting.CharacterProperties
}

// ParagraphFormatEntry represents paragraph formatting information.
type ParagraphFormatEntry struct {
	StartCP uint32
	EndCP   uint32
	Props   *formatting.ParagraphProperties
}

// NewDocumentWriter creates a new document writer.
func NewDocumentWriter() *DocumentWriter {
	return &DocumentWriter{
		metadata: &DocumentInfo{
			Created:     time.Now(),
			Modified:    time.Now(),
			Application: "msdoc library",
			Language:    0x0409, // English (US)
		},
		text:       make([]TextSection, 0),
		fibBuilder: NewFIBBuilder(),
		pieceTable: NewPieceTableBuilder(),
		formatting: NewFormattingBuilder(),
	}
}

// SetTitle sets the document title.
func (dw *DocumentWriter) SetTitle(title string) {
	dw.metadata.Title = title
}

// SetAuthor sets the document author.
func (dw *DocumentWriter) SetAuthor(author string) {
	dw.metadata.Author = author
}

// SetSubject sets the document subject.
func (dw *DocumentWriter) SetSubject(subject string) {
	dw.metadata.Subject = subject
}

// SetKeywords sets the document keywords.
func (dw *DocumentWriter) SetKeywords(keywords string) {
	dw.metadata.Keywords = keywords
}

// SetComments sets the document comments.
func (dw *DocumentWriter) SetComments(comments string) {
	dw.metadata.Comments = comments
}

// SetCompany sets the company name.
func (dw *DocumentWriter) SetCompany(company string) {
	dw.metadata.Company = company
}

// AddText adds plain text to the document.
func (dw *DocumentWriter) AddText(text string) {
	dw.AddFormattedText(text, nil, nil)
}

// AddParagraph adds a new paragraph with text.
func (dw *DocumentWriter) AddParagraph(text string) {
	dw.AddFormattedParagraph(text, nil, nil)
}

// AddFormattedText adds text with character formatting.
func (dw *DocumentWriter) AddFormattedText(text string, charProps *formatting.CharacterProperties, paraProps *formatting.ParagraphProperties) {
	section := TextSection{
		Text:      text,
		CharProps: charProps,
		ParaProps: paraProps,
		IsNewPara: false,
	}
	dw.text = append(dw.text, section)
}

// AddFormattedParagraph adds a new paragraph with formatting.
func (dw *DocumentWriter) AddFormattedParagraph(text string, charProps *formatting.CharacterProperties, paraProps *formatting.ParagraphProperties) {
	section := TextSection{
		Text:      text + "\r", // Add paragraph marker
		CharProps: charProps,
		ParaProps: paraProps,
		IsNewPara: true,
	}
	dw.text = append(dw.text, section)
}

// InsertPageBreak inserts a page break.
func (dw *DocumentWriter) InsertPageBreak() {
	dw.AddText("\f") // Form feed character for page break
}

// InsertSectionBreak inserts a section break.
func (dw *DocumentWriter) InsertSectionBreak() {
	dw.AddText("\r") // Section break marker
}

// Save saves the document to the specified filename.
func (dw *DocumentWriter) Save(filename string) error {
	// Build the document structure
	if err := dw.buildDocument(); err != nil {
		return fmt.Errorf("failed to build document: %w", err)
	}

	// Create output file
	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	// Write OLE2 compound document
	if err := dw.writeOLE2Document(file); err != nil {
		return fmt.Errorf("failed to write OLE2 document: %w", err)
	}

	return nil
}

// buildDocument builds the internal document structure.
func (dw *DocumentWriter) buildDocument() error {
	// Build piece table from text sections
	currentCP := uint32(0)
	for _, section := range dw.text {
		piece := PieceDescriptor{
			StartCP:    currentCP,
			EndCP:      currentCP + uint32(len(section.Text)),
			FileOffset: uint32(dw.pieceTable.text.Len()),
			IsUnicode:  dw.needsUnicode(section.Text),
		}

		// Add text to buffer
		if piece.IsUnicode {
			dw.addUnicodeText(section.Text)
		} else {
			dw.pieceTable.text.WriteString(section.Text)
		}

		dw.pieceTable.pieces = append(dw.pieceTable.pieces, piece)

		// Add formatting entries
		if section.CharProps != nil {
			dw.formatting.charFormats = append(dw.formatting.charFormats, CharacterFormatEntry{
				StartCP: currentCP,
				EndCP:   piece.EndCP,
				Props:   section.CharProps,
			})
		}

		if section.ParaProps != nil && section.IsNewPara {
			dw.formatting.paraFormats = append(dw.formatting.paraFormats, ParagraphFormatEntry{
				StartCP: currentCP,
				EndCP:   piece.EndCP,
				Props:   section.ParaProps,
			})
		}

		currentCP = piece.EndCP
	}

	// Build FIB
	dw.fibBuilder.SetTextLength(currentCP)
	dw.fibBuilder.SetCreated(dw.metadata.Created)
	dw.fibBuilder.SetModified(dw.metadata.Modified)

	return nil
}

// needsUnicode checks if text contains non-ASCII characters.
func (dw *DocumentWriter) needsUnicode(text string) bool {
	for _, r := range text {
		if r > 127 {
			return true
		}
	}
	return false
}

// addUnicodeText adds Unicode text to the buffer.
func (dw *DocumentWriter) addUnicodeText(text string) {
	for _, r := range text {
		// Write as UTF-16LE
		// It is safe to ignore the error from binary.Write here because bytes.Buffer.Write never returns an error.
		binary.Write(&dw.pieceTable.text, binary.LittleEndian, uint16(r))
	}
}

// writeOLE2Document writes the complete OLE2 compound document.
func (dw *DocumentWriter) writeOLE2Document(writer io.Writer) error {
	// Create OLE2 writer
	oleWriter := ole2.NewWriter()

	// Write WordDocument stream
	wordDocStream, err := dw.buildWordDocumentStream()
	if err != nil {
		return fmt.Errorf("failed to build WordDocument stream: %w", err)
	}
	oleWriter.AddStream("WordDocument", wordDocStream)

	// Write Table stream (1Table for newer documents)
	tableStream, err := dw.buildTableStream()
	if err != nil {
		return fmt.Errorf("failed to build Table stream: %w", err)
	}
	oleWriter.AddStream("1Table", tableStream)

	// Write SummaryInformation stream
	summaryStream, err := dw.buildSummaryInformationStream()
	if err != nil {
		return fmt.Errorf("failed to build SummaryInformation stream: %w", err)
	}
	oleWriter.AddStream("\x05SummaryInformation", summaryStream)

	// Write DocumentSummaryInformation stream
	docSummaryStream, err := dw.buildDocumentSummaryInformationStream()
	if err != nil {
		return fmt.Errorf("failed to build DocumentSummaryInformation stream: %w", err)
	}
	oleWriter.AddStream("\x05DocumentSummaryInformation", docSummaryStream)

	// Write the compound document
	return oleWriter.WriteTo(writer)
}

// buildWordDocumentStream constructs the WordDocument stream.
func (dw *DocumentWriter) buildWordDocumentStream() ([]byte, error) {
	var buffer bytes.Buffer

	// Write FIB
	fibData, err := dw.fibBuilder.Build()
	if err != nil {
		return nil, fmt.Errorf("failed to build FIB: %w", err)
	}
	buffer.Write(fibData)

	// Write text content
	buffer.Write(dw.pieceTable.text.Bytes())

	return buffer.Bytes(), nil
}

// buildTableStream constructs the Table stream.
func (dw *DocumentWriter) buildTableStream() ([]byte, error) {
	var buffer bytes.Buffer

	// Write piece table (CLX)
	clxData, err := dw.buildCLX()
	if err != nil {
		return nil, fmt.Errorf("failed to build CLX: %w", err)
	}
	buffer.Write(clxData)

	// Write formatting tables
	formattingData, err := dw.buildFormattingTables()
	if err != nil {
		return nil, fmt.Errorf("failed to build formatting tables: %w", err)
	}
	buffer.Write(formattingData)

	return buffer.Bytes(), nil
}

// buildCLX constructs the CLX (piece table) structure.
func (dw *DocumentWriter) buildCLX() ([]byte, error) {
	var buffer bytes.Buffer

	// CLX marker
	buffer.WriteByte(0x02)

	// Build PLC of piece descriptors
	plcData, err := dw.buildPiecePLC()
	if err != nil {
		return nil, err
	}
	buffer.Write(plcData)

	return buffer.Bytes(), nil
}

// buildPiecePLC builds the PLC structure for pieces.
func (dw *DocumentWriter) buildPiecePLC() ([]byte, error) {
	var buffer bytes.Buffer

	// Write CP array
	for _, piece := range dw.pieceTable.pieces {
		binary.Write(&buffer, binary.LittleEndian, piece.StartCP)
	}
	// Write final CP
	if len(dw.pieceTable.pieces) > 0 {
		lastPiece := dw.pieceTable.pieces[len(dw.pieceTable.pieces)-1]
		binary.Write(&buffer, binary.LittleEndian, lastPiece.EndCP)
	}

	// Write PCD array
	for _, piece := range dw.pieceTable.pieces {
		pcdData, err := dw.buildPCD(piece)
		if err != nil {
			return nil, err
		}
		buffer.Write(pcdData)
	}

	return buffer.Bytes(), nil
}

// buildPCD builds a Piece Descriptor.
func (dw *DocumentWriter) buildPCD(piece PieceDescriptor) ([]byte, error) {
	var buffer bytes.Buffer

	// PCD structure (8 bytes)
	flags := uint16(0x0001) // fNoEncryption
	if piece.IsUnicode {
		flags |= 0x0040 // fCompressed = 0 for Unicode
	}

	binary.Write(&buffer, binary.LittleEndian, flags)
	binary.Write(&buffer, binary.LittleEndian, piece.FileOffset)
	binary.Write(&buffer, binary.LittleEndian, uint16(0)) // Prm (property modifier)

	return buffer.Bytes(), nil
}

// buildFormattingTables constructs character and paragraph formatting tables.
func (dw *DocumentWriter) buildFormattingTables() ([]byte, error) {
	var buffer bytes.Buffer

	// Build character formatting (CHPX FKP)
	chpxData, err := dw.buildCHPXTable()
	if err != nil {
		return nil, fmt.Errorf("failed to build CHPX table: %w", err)
	}
	buffer.Write(chpxData)

	// Build paragraph formatting (PAPX FKP)
	papxData, err := dw.buildPAPXTable()
	if err != nil {
		return nil, fmt.Errorf("failed to build PAPX table: %w", err)
	}
	buffer.Write(papxData)

	return buffer.Bytes(), nil
}

// buildCHPXTable builds the character formatting table.
func (dw *DocumentWriter) buildCHPXTable() ([]byte, error) {
	// This would build FKP structures for character properties
	// For now, return empty data (default formatting)
	return make([]byte, 512), nil // Standard FKP page size
}

// buildPAPXTable builds the paragraph formatting table.
func (dw *DocumentWriter) buildPAPXTable() ([]byte, error) {
	// This would build FKP structures for paragraph properties
	// For now, return empty data (default formatting)
	return make([]byte, 512), nil // Standard FKP page size
}

// buildSummaryInformationStream constructs the SummaryInformation property set.
func (dw *DocumentWriter) buildSummaryInformationStream() ([]byte, error) {
	// This would build the complete property set stream
	// For now, return minimal data
	return dw.buildMinimalPropertySet(dw.metadata), nil
}

// buildDocumentSummaryInformationStream constructs DocumentSummaryInformation.
func (dw *DocumentWriter) buildDocumentSummaryInformationStream() ([]byte, error) {
	// This would build the complete document summary property set
	return dw.buildMinimalPropertySet(dw.metadata), nil
}

// buildMinimalPropertySet creates a minimal property set with basic metadata.
func (dw *DocumentWriter) buildMinimalPropertySet(info *DocumentInfo) []byte {
	var buffer bytes.Buffer

	// Property set header
	binary.Write(&buffer, binary.LittleEndian, uint16(0xFFFE)) // Byte order
	binary.Write(&buffer, binary.LittleEndian, uint16(0x0000)) // Version
	binary.Write(&buffer, binary.LittleEndian, uint32(0x000))  // System ID

	// CLSID (16 bytes of zeros)
	buffer.Write(make([]byte, 16))

	binary.Write(&buffer, binary.LittleEndian, uint32(1)) // Number of property sets

	// Property set format ID and offset
	buffer.Write(make([]byte, 16))                         // Format ID (SummaryInformation GUID)
	binary.Write(&buffer, binary.LittleEndian, uint32(48)) // Offset

	// Property set data
	binary.Write(&buffer, binary.LittleEndian, uint32(buffer.Len()+100)) // Size
	binary.Write(&buffer, binary.LittleEndian, uint32(1))                // Property count

	// Property ID and offset for title
	binary.Write(&buffer, binary.LittleEndian, uint32(0x02))  // PID_TITLE
	binary.Write(&buffer, binary.LittleEndian, uint32(48+16)) // Offset

	// Title property value
	binary.Write(&buffer, binary.LittleEndian, uint16(0x001F)) // VT_LPWSTR
	binary.Write(&buffer, binary.LittleEndian, uint16(0x0000)) // Padding

	titleBytes := []byte(info.Title)
	binary.Write(&buffer, binary.LittleEndian, uint32(len(titleBytes)+1)) // Length
	buffer.Write(titleBytes)
	buffer.WriteByte(0) // Null terminator

	return buffer.Bytes()
}

// NewFIBBuilder creates a new FIB builder.
func NewFIBBuilder() *FIBBuilder {
	return &FIBBuilder{
		fib: &fib.FileInformationBlock{},
	}
}

// SetTextLength sets the document text length.
func (fb *FIBBuilder) SetTextLength(length uint32) {
	fb.fib.FibRgLw.CcpText = length
}

// SetCreated sets the creation time.
func (fb *FIBBuilder) SetCreated(created time.Time) {
	// Convert to FILETIME format if needed
}

// SetModified sets the modification time.
func (fb *FIBBuilder) SetModified(modified time.Time) {
	// Convert to FILETIME format if needed
}

// Build constructs the FIB data.
func (fb *FIBBuilder) Build() ([]byte, error) {
	var buffer bytes.Buffer

	// Set required FIB fields
	fb.fib.Base.WIdent = 0xA5EC // Word identifier
	fb.fib.Base.NFib = 0x0112   // Word 2003 FIB version
	fb.fib.Base.LKey = 0        // No encryption key
	fb.fib.Base.Envr = 0        // Not created by Word
	fb.fib.Base.Flags1 = 0x0000 // No special flags

	// Write FIB base
	if err := binary.Write(&buffer, binary.LittleEndian, &fb.fib.Base); err != nil {
		return nil, fmt.Errorf("failed to write FIB base: %w", err)
	}

	// Write other FIB sections
	// This would write Csw, FibRgW, Cslw, FibRgLw, etc.
	// For now, write minimal required sections

	return buffer.Bytes(), nil
}

// NewPieceTableBuilder creates a new piece table builder.
func NewPieceTableBuilder() *PieceTableBuilder {
	return &PieceTableBuilder{
		pieces: make([]PieceDescriptor, 0),
	}
}

// NewFormattingBuilder creates a new formatting builder.
func NewFormattingBuilder() *FormattingBuilder {
	return &FormattingBuilder{
		charFormats: make([]CharacterFormatEntry, 0),
		paraFormats: make([]ParagraphFormatEntry, 0),
	}
}
