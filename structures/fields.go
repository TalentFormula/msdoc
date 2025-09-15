package structures

import (
	"fmt"
	"strings"
)

// Field represents a field in a Word document (used for hyperlinks, etc.)
type Field struct {
	Start      CP     // Character position where field starts
	End        CP     // Character position where field ends
	FieldType  byte   // Field type (19h for HYPERLINK)
	FieldCode  string // The field code (e.g., "HYPERLINK \"url\"")
	DisplayText string // The display text for the field
}

// HyperlinkField represents a parsed hyperlink field
type HyperlinkField struct {
	URL         string
	DisplayText string
	Start       CP
	End         CP
}

// FieldPLC represents a field PLC (Piece Location Collection)
type FieldPLC struct {
	*PLC
}

// ParseFieldPLC creates a FieldPLC from raw bytes
func ParseFieldPLC(data []byte) (*FieldPLC, error) {
	// Field PLCs have 2-byte data elements (FLD structure)
	plc, err := ParsePLC(data, 2)
	if err != nil {
		return nil, fmt.Errorf("failed to parse field PLC: %w", err)
	}

	return &FieldPLC{PLC: plc}, nil
}

// GetFields extracts all fields from the PLC
func (fplc *FieldPLC) GetFields() ([]*Field, error) {
	if fplc.PLC == nil {
		return nil, fmt.Errorf("no PLC data available")
	}

	fields := make([]*Field, 0)

	// Fields come in pairs: field start and field end
	for i := 0; i < fplc.Count(); i += 2 {
		if i+1 >= fplc.Count() {
			break // Incomplete pair
		}

		startCP, endCP, err := fplc.GetRange(i)
		if err != nil {
			continue
		}

		fieldData, err := fplc.GetDataAt(i)
		if err != nil {
			continue
		}

		// Parse FLD structure (2 bytes)
		if len(fieldData) < 2 {
			continue
		}

		// Field type is in the lower 5 bits of first byte
		fieldType := fieldData[0] & 0x1F

		field := &Field{
			Start:     startCP,
			End:       endCP,
			FieldType: fieldType,
		}

		fields = append(fields, field)
	}

	return fields, nil
}

// ExtractHyperlinks extracts hyperlink information from text and field data
func ExtractHyperlinks(text string, fields []*Field) ([]*HyperlinkField, error) {
	hyperlinks := make([]*HyperlinkField, 0)

	for _, field := range fields {
		// Field type 58 (0x3A) is HYPERLINK in Word
		if field.FieldType == 0x13 || field.FieldType == 0x3A {
			// This is a hyperlink field
			startPos := int(field.Start)
			endPos := int(field.End)

			if startPos >= 0 && endPos >= startPos && endPos <= len(text) {
				fieldText := text[startPos:endPos]

				// Parse hyperlink field code
				url, displayText := parseHyperlinkField(fieldText)
				if url != "" {
					hyperlink := &HyperlinkField{
						URL:         url,
						DisplayText: displayText,
						Start:       field.Start,
						End:         field.End,
					}
					hyperlinks = append(hyperlinks, hyperlink)
				}
			}
		}
	}

	return hyperlinks, nil
}

// parseHyperlinkField parses a hyperlink field code to extract URL and display text
func parseHyperlinkField(fieldText string) (url, displayText string) {
	// Field codes often contain special characters, try to extract URL
	// Field format is typically: HYPERLINK "url" \o "tooltip" displaytext
	
	// Look for HYPERLINK keyword
	if !strings.Contains(strings.ToUpper(fieldText), "HYPERLINK") {
		return "", ""
	}

	// Simple extraction - look for quoted URLs
	parts := strings.Fields(fieldText)
	for i, part := range parts {
		if strings.ToUpper(part) == "HYPERLINK" && i+1 < len(parts) {
			urlPart := parts[i+1]
			// Remove quotes
			url = strings.Trim(urlPart, "\"")
			
			// The rest might be display text
			if i+2 < len(parts) {
				displayText = strings.Join(parts[i+2:], " ")
				displayText = strings.Trim(displayText, "\"")
			}
			break
		}
	}

	return url, displayText
}

// FormatAsMarkdown formats hyperlinks as markdown [text](url)
func (hl *HyperlinkField) FormatAsMarkdown() string {
	if hl.DisplayText != "" {
		return fmt.Sprintf("[%s](%s)", hl.DisplayText, hl.URL)
	}
	return fmt.Sprintf("[%s](%s)", hl.URL, hl.URL)
}