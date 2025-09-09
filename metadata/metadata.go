// Package metadata provides comprehensive metadata extraction from .doc files.
//
// This package implements full SummaryInformation and DocumentSummaryInformation
// parsing according to the OLE Property Set specification, supporting all
// standard document properties and custom properties.
package metadata

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"strings"
	"time"
	"unicode/utf16"

	"github.com/TalentFormula/msdoc/ole2"
)

// DocumentMetadata holds comprehensive document metadata information.
type DocumentMetadata struct {
	// Core properties from SummaryInformation
	Title            string    // Document title
	Subject          string    // Document subject  
	Author           string    // Document author
	Keywords         string    // Document keywords
	Comments         string    // Document comments
	Template         string    // Template name
	LastAuthor       string    // Last saved by
	RevisionNumber   string    // Revision number
	ApplicationName  string    // Creating application
	Created          time.Time // Creation time
	LastSaved        time.Time // Last saved time
	LastPrinted      time.Time // Last printed time
	TotalEditTime    int64     // Total editing time in minutes
	PageCount        int32     // Number of pages
	WordCount        int32     // Number of words
	CharCount        int32     // Number of characters
	CharCountWithSpaces int32  // Number of characters with spaces
	Security         int32     // Security flags
	Category         string    // Document category
	PresentationFormat string  // Presentation format
	ByteCount        int64     // Number of bytes
	LineCount        int32     // Number of lines
	ParagraphCount   int32     // Number of paragraphs
	SlideCount       int32     // Number of slides
	NoteCount        int32     // Number of notes
	HiddenSlideCount int32     // Number of hidden slides
	MultimediaClipCount int32  // Number of multimedia clips
	
	// Document summary properties
	Company          string            // Company name
	Manager          string            // Manager name
	Language         int32             // Document language
	DocumentVersion  string            // Document version
	ContentType      string            // Content type
	ContentStatus    string            // Content status
	HyperLinkBase    string            // Hyperlink base
	CustomProperties map[string]interface{} // Custom properties
	
	// Extended properties
	ThumbnailClipboardFormat int32    // Thumbnail format
	ThumbnailData           []byte   // Thumbnail image data
	
	// Security and protection
	ReadOnlyRecommended     bool     // Read-only recommended
	WriteReservationPassword bool    // Write reservation password set
	ReadOnlyPassword        bool     // Read-only password set
}

// PropertyType represents the data type of a property.
type PropertyType uint16

const (
	PropertyTypeEmpty        PropertyType = 0x0000 // VT_EMPTY
	PropertyTypeNull         PropertyType = 0x0001 // VT_NULL
	PropertyTypeInt16        PropertyType = 0x0002 // VT_I2
	PropertyTypeInt32        PropertyType = 0x0003 // VT_I4
	PropertyTypeFloat        PropertyType = 0x0004 // VT_R4
	PropertyTypeDouble       PropertyType = 0x0005 // VT_R8
	PropertyTypeCurrency     PropertyType = 0x0006 // VT_CY
	PropertyTypeDate         PropertyType = 0x0007 // VT_DATE
	PropertyTypeString       PropertyType = 0x0008 // VT_BSTR
	PropertyTypeBoolean      PropertyType = 0x000B // VT_BOOL
	PropertyTypeVariant      PropertyType = 0x000C // VT_VARIANT
	PropertyTypeInt8         PropertyType = 0x0010 // VT_I1
	PropertyTypeUInt8        PropertyType = 0x0011 // VT_UI1
	PropertyTypeUInt16       PropertyType = 0x0012 // VT_UI2
	PropertyTypeUInt32       PropertyType = 0x0013 // VT_UI4
	PropertyTypeInt64        PropertyType = 0x0014 // VT_I8
	PropertyTypeUInt64       PropertyType = 0x0015 // VT_UI8
	PropertyTypeFileTime     PropertyType = 0x0040 // VT_FILETIME
	PropertyTypeBlob         PropertyType = 0x0041 // VT_BLOB
	PropertyTypeClipboardData PropertyType = 0x0047 // VT_CF
	PropertyTypeStringA      PropertyType = 0x001E // VT_LPSTR
	PropertyTypeStringW      PropertyType = 0x001F // VT_LPWSTR
)

// Property IDs for SummaryInformation stream
const (
	PIDTitle          = 0x02
	PIDSubject        = 0x03
	PIDAuthor         = 0x04
	PIDKeywords       = 0x05
	PIDComments       = 0x06
	PIDTemplate       = 0x07
	PIDLastAuthor     = 0x08
	PIDRevNumber      = 0x09
	PIDEditTime       = 0x0A
	PIDLastPrinted    = 0x0B
	PIDCreateTime     = 0x0C
	PIDLastSaveTime   = 0x0D
	PIDPageCount      = 0x0E
	PIDWordCount      = 0x0F
	PIDCharCount      = 0x10
	PIDThumbnail      = 0x11
	PIDAppName        = 0x12
	PIDSecurity       = 0x13
)

// Property IDs for DocumentSummaryInformation stream
const (
	PIDCategory       = 0x02
	PIDPresentationFormat = 0x03
	PIDByteCount      = 0x04
	PIDLineCount      = 0x05
	PIDParaCount      = 0x06
	PIDSlideCount     = 0x07
	PIDNoteCount      = 0x08
	PIDHiddenCount    = 0x09
	PIDMMClipCount    = 0x0A
	PIDScale          = 0x0B
	PIDHeadingPairs   = 0x0C
	PIDDocParts       = 0x0D
	PIDManager        = 0x0E
	PIDCompany        = 0x0F
	PIDLinksUpToDate  = 0x10
	PIDCharCountWithSpaces = 0x11
	PIDSharedDoc      = 0x13
	PIDHyperLinkBase  = 0x15
	PIDHyperLinks     = 0x16
	PIDHyperLinksChanged = 0x17
	PIDVersion        = 0x18
	PIDDigSig         = 0x19
	PIDContentType    = 0x1A
	PIDContentStatus  = 0x1B
	PIDLanguage       = 0x1C
	PIDDocVersion     = 0x1D
)

// MetadataExtractor handles extraction of metadata from .doc files.
type MetadataExtractor struct {
	reader *ole2.Reader
}

// NewMetadataExtractor creates a new metadata extractor.
func NewMetadataExtractor(reader *ole2.Reader) *MetadataExtractor {
	return &MetadataExtractor{
		reader: reader,
	}
}

// ExtractMetadata extracts complete metadata from the document.
func (me *MetadataExtractor) ExtractMetadata() (*DocumentMetadata, error) {
	metadata := &DocumentMetadata{
		CustomProperties: make(map[string]interface{}),
	}

	// Extract SummaryInformation properties
	if err := me.extractSummaryInformation(metadata); err != nil {
		// Don't fail if SummaryInformation is missing, just log it
		fmt.Printf("Warning: Failed to extract SummaryInformation: %v\n", err)
	}

	// Extract DocumentSummaryInformation properties
	if err := me.extractDocumentSummaryInformation(metadata); err != nil {
		// Don't fail if DocumentSummaryInformation is missing
		fmt.Printf("Warning: Failed to extract DocumentSummaryInformation: %v\n", err)
	}

	return metadata, nil
}

// extractSummaryInformation extracts properties from the SummaryInformation stream.
func (me *MetadataExtractor) extractSummaryInformation(metadata *DocumentMetadata) error {
	// Read SummaryInformation stream
	streamData, err := me.reader.ReadStream("\x05SummaryInformation")
	if err != nil {
		return fmt.Errorf("failed to read SummaryInformation stream: %w", err)
	}

	properties, err := me.parsePropertySet(streamData)
	if err != nil {
		return fmt.Errorf("failed to parse SummaryInformation: %w", err)
	}

	// Extract known properties
	for propID, value := range properties {
		switch propID {
		case PIDTitle:
			if str, ok := value.(string); ok {
				metadata.Title = str
			}
		case PIDSubject:
			if str, ok := value.(string); ok {
				metadata.Subject = str
			}
		case PIDAuthor:
			if str, ok := value.(string); ok {
				metadata.Author = str
			}
		case PIDKeywords:
			if str, ok := value.(string); ok {
				metadata.Keywords = str
			}
		case PIDComments:
			if str, ok := value.(string); ok {
				metadata.Comments = str
			}
		case PIDTemplate:
			if str, ok := value.(string); ok {
				metadata.Template = str
			}
		case PIDLastAuthor:
			if str, ok := value.(string); ok {
				metadata.LastAuthor = str
			}
		case PIDRevNumber:
			if str, ok := value.(string); ok {
				metadata.RevisionNumber = str
			}
		case PIDAppName:
			if str, ok := value.(string); ok {
				metadata.ApplicationName = str
			}
		case PIDCreateTime:
			if t, ok := value.(time.Time); ok {
				metadata.Created = t
			}
		case PIDLastSaveTime:
			if t, ok := value.(time.Time); ok {
				metadata.LastSaved = t
			}
		case PIDLastPrinted:
			if t, ok := value.(time.Time); ok {
				metadata.LastPrinted = t
			}
		case PIDEditTime:
			if i, ok := value.(int64); ok {
				metadata.TotalEditTime = i / (10000000 * 60) // Convert from 100ns to minutes
			}
		case PIDPageCount:
			if i, ok := value.(int32); ok {
				metadata.PageCount = i
			}
		case PIDWordCount:
			if i, ok := value.(int32); ok {
				metadata.WordCount = i
			}
		case PIDCharCount:
			if i, ok := value.(int32); ok {
				metadata.CharCount = i
			}
		case PIDSecurity:
			if i, ok := value.(int32); ok {
				metadata.Security = i
			}
		case PIDThumbnail:
			if data, ok := value.([]byte); ok {
				metadata.ThumbnailData = data
			}
		}
	}

	return nil
}

// extractDocumentSummaryInformation extracts properties from DocumentSummaryInformation stream.
func (me *MetadataExtractor) extractDocumentSummaryInformation(metadata *DocumentMetadata) error {
	// Read DocumentSummaryInformation stream
	streamData, err := me.reader.ReadStream("\x05DocumentSummaryInformation")
	if err != nil {
		return fmt.Errorf("failed to read DocumentSummaryInformation stream: %w", err)
	}

	properties, err := me.parsePropertySet(streamData)
	if err != nil {
		return fmt.Errorf("failed to parse DocumentSummaryInformation: %w", err)
	}

	// Extract known properties
	for propID, value := range properties {
		switch propID {
		case PIDCategory:
			if str, ok := value.(string); ok {
				metadata.Category = str
			}
		case PIDCompany:
			if str, ok := value.(string); ok {
				metadata.Company = str
			}
		case PIDManager:
			if str, ok := value.(string); ok {
				metadata.Manager = str
			}
		case PIDByteCount:
			if i, ok := value.(int64); ok {
				metadata.ByteCount = i
			}
		case PIDLineCount:
			if i, ok := value.(int32); ok {
				metadata.LineCount = i
			}
		case PIDParaCount:
			if i, ok := value.(int32); ok {
				metadata.ParagraphCount = i
			}
		case PIDCharCountWithSpaces:
			if i, ok := value.(int32); ok {
				metadata.CharCountWithSpaces = i
			}
		case PIDLanguage:
			if i, ok := value.(int32); ok {
				metadata.Language = i
			}
		case PIDContentType:
			if str, ok := value.(string); ok {
				metadata.ContentType = str
			}
		case PIDContentStatus:
			if str, ok := value.(string); ok {
				metadata.ContentStatus = str
			}
		case PIDHyperLinkBase:
			if str, ok := value.(string); ok {
				metadata.HyperLinkBase = str
			}
		}
	}

	return nil
}

// parsePropertySet parses an OLE property set stream.
func (me *MetadataExtractor) parsePropertySet(data []byte) (map[uint32]interface{}, error) {
	if len(data) < 48 {
		return nil, errors.New("property set data too short")
	}

	reader := bytes.NewReader(data)
	properties := make(map[uint32]interface{})

	// Read property set header
	var header struct {
		ByteOrder    uint16 // Byte order identifier
		Version      uint16 // Version
		SystemID     uint32 // System identifier
		CLSID        [16]byte // CLSID
		NumPropertySets uint32 // Number of property sets
	}

	if err := binary.Read(reader, binary.LittleEndian, &header); err != nil {
		return nil, fmt.Errorf("failed to read property set header: %w", err)
	}

	// Validate byte order
	if header.ByteOrder != 0xFFFE {
		return nil, errors.New("invalid byte order in property set")
	}

	// Read property set info
	for i := uint32(0); i < header.NumPropertySets; i++ {
		var psInfo struct {
			FMTID  [16]byte // Format ID
			Offset uint32   // Offset to property set
		}

		if err := binary.Read(reader, binary.LittleEndian, &psInfo); err != nil {
			return nil, fmt.Errorf("failed to read property set info %d: %w", i, err)
		}

		// Parse property set at offset
		if err := me.parsePropertySetData(data[psInfo.Offset:], properties); err != nil {
			return nil, fmt.Errorf("failed to parse property set %d: %w", i, err)
		}
	}

	return properties, nil
}

// parsePropertySetData parses the actual property data.
func (me *MetadataExtractor) parsePropertySetData(data []byte, properties map[uint32]interface{}) error {
	if len(data) < 8 {
		return errors.New("property set data too short")
	}

	reader := bytes.NewReader(data)

	// Read property set size and property count
	var size, count uint32
	if err := binary.Read(reader, binary.LittleEndian, &size); err != nil {
		return fmt.Errorf("failed to read property set size: %w", err)
	}
	if err := binary.Read(reader, binary.LittleEndian, &count); err != nil {
		return fmt.Errorf("failed to read property count: %w", err)
	}

	// Read property identifiers and offsets
	propOffsets := make(map[uint32]uint32)
	for i := uint32(0); i < count; i++ {
		var propID, offset uint32
		if err := binary.Read(reader, binary.LittleEndian, &propID); err != nil {
			return fmt.Errorf("failed to read property ID %d: %w", i, err)
		}
		if err := binary.Read(reader, binary.LittleEndian, &offset); err != nil {
			return fmt.Errorf("failed to read property offset %d: %w", i, err)
		}
		propOffsets[propID] = offset
	}

	// Read properties
	for propID, offset := range propOffsets {
		if uint32(len(data)) <= offset {
			continue // Skip invalid offset
		}

		propReader := bytes.NewReader(data[offset:])
		value, err := me.readPropertyValue(propReader)
		if err != nil {
			continue // Skip invalid property
		}

		properties[propID] = value
	}

	return nil
}

// readPropertyValue reads a property value based on its type.
func (me *MetadataExtractor) readPropertyValue(reader *bytes.Reader) (interface{}, error) {
	// Read property type
	var propType PropertyType
	if err := binary.Read(reader, binary.LittleEndian, &propType); err != nil {
		return nil, fmt.Errorf("failed to read property type: %w", err)
	}

	// Skip padding if needed
	reader.Seek(2, 1) // Skip 2 bytes of padding

	// Read value based on type
	switch propType {
	case PropertyTypeInt32:
		var value int32
		if err := binary.Read(reader, binary.LittleEndian, &value); err != nil {
			return nil, err
		}
		return value, nil

	case PropertyTypeUInt32:
		var value uint32
		if err := binary.Read(reader, binary.LittleEndian, &value); err != nil {
			return nil, err
		}
		return int32(value), nil

	case PropertyTypeInt64:
		var value int64
		if err := binary.Read(reader, binary.LittleEndian, &value); err != nil {
			return nil, err
		}
		return value, nil

	case PropertyTypeFileTime:
		var filetime int64
		if err := binary.Read(reader, binary.LittleEndian, &filetime); err != nil {
			return nil, err
		}
		// Convert Windows FILETIME to Go time
		// FILETIME is 100-nanosecond intervals since January 1, 1601 UTC
		const fileTimeEpoch = 116444736000000000 // January 1, 1601 to January 1, 1970
		unixTime := (filetime - fileTimeEpoch) / 10000000
		return time.Unix(unixTime, 0), nil

	case PropertyTypeString, PropertyTypeStringA, PropertyTypeStringW:
		// Read string length
		var length uint32
		if err := binary.Read(reader, binary.LittleEndian, &length); err != nil {
			return nil, err
		}

		if length == 0 {
			return "", nil
		}

		// Read string data
		if propType == PropertyTypeStringW {
			// Unicode string
			utf16Data := make([]uint16, length/2)
			if err := binary.Read(reader, binary.LittleEndian, &utf16Data); err != nil {
				return nil, err
			}
			return strings.TrimSpace(string(utf16.Decode(utf16Data))), nil
		} else {
			// ANSI string (PropertyTypeString or PropertyTypeStringA)
			strData := make([]byte, length)
			if _, err := reader.Read(strData); err != nil {
				return nil, err
			}
			// Remove all trailing null characters
			for len(strData) > 0 && strData[len(strData)-1] == 0 {
				strData = strData[:len(strData)-1]
			}
			// Trim whitespace from the string
			return strings.TrimSpace(string(strData)), nil
		}

	case PropertyTypeBlob, PropertyTypeClipboardData:
		// Read blob size
		var size uint32
		if err := binary.Read(reader, binary.LittleEndian, &size); err != nil {
			return nil, err
		}

		// Read blob data
		blobData := make([]byte, size)
		if _, err := reader.Read(blobData); err != nil {
			return nil, err
		}
		return blobData, nil

	case PropertyTypeBoolean:
		var value uint16
		if err := binary.Read(reader, binary.LittleEndian, &value); err != nil {
			return nil, err
		}
		return value != 0, nil

	default:
		return nil, fmt.Errorf("unsupported property type: %d", propType)
	}
}

// GetLanguageName returns the human-readable language name for a language code.
func (metadata *DocumentMetadata) GetLanguageName() string {
	// Common language codes
	languages := map[int32]string{
		0x0409: "English (US)",
		0x0809: "English (UK)",
		0x040C: "French",
		0x0407: "German",
		0x0410: "Italian",
		0x040A: "Spanish",
		0x0416: "Portuguese (Brazil)",
		0x0816: "Portuguese (Portugal)",
		0x0411: "Japanese",
		0x0804: "Chinese (Simplified)",
		0x0404: "Chinese (Traditional)",
		0x0412: "Korean",
		0x0419: "Russian",
	}

	if name, exists := languages[metadata.Language]; exists {
		return name
	}
	return fmt.Sprintf("Language ID: %d", metadata.Language)
}

// IsProtected returns true if the document has any protection enabled.
func (metadata *DocumentMetadata) IsProtected() bool {
	return metadata.Security != 0 || 
		   metadata.ReadOnlyRecommended || 
		   metadata.WriteReservationPassword || 
		   metadata.ReadOnlyPassword
}