package streams

import (
	"fmt"

	"github.com/TalentFormula/msdoc/fib"
	"github.com/TalentFormula/msdoc/structures"
)

// WordDocumentStream represents the main document stream containing text and the FIB.
type WordDocumentStream struct {
	Data []byte
	FIB  *fib.FileInformationBlock
}

// NewWordDocumentStream creates a new WordDocument stream processor.
func NewWordDocumentStream(data []byte) (*WordDocumentStream, error) {
	// Parse the FIB from the beginning of the stream
	parsedFIB, err := fib.ParseFIB(data)
	if err != nil {
		return nil, err
	}

	return &WordDocumentStream{
		Data: data,
		FIB:  parsedFIB,
	}, nil
}

// GetTextRange extracts text from the specified character range.
func (wds *WordDocumentStream) GetTextRange(start, end structures.CP) ([]byte, error) {
	if start >= end {
		return nil, nil
	}

	// Validate character positions
	if !start.IsValid() || !end.IsValid() {
		return nil, structures.ErrInvalidCP
	}

	// Calculate the range size
	size := start.Distance(end)
	if size == 0 {
		return nil, nil
	}

	// Extract the text from the stream starting at fcMin + start
	// Note: This is a simplified implementation. Real DOC files use piece tables.
	textStart := wds.FIB.FibRgLw.CbMac + uint32(start)
	if textStart+size > uint32(len(wds.Data)) {
		return nil, ErrStreamTooSmall
	}

	result := make([]byte, size)
	copy(result, wds.Data[textStart:textStart+size])
	return result, nil
}

// GetMainTextLength returns the length of the main document text in characters.
func (wds *WordDocumentStream) GetMainTextLength() uint32 {
	return wds.FIB.FibRgLw.CcpText
}

// GetTotalTextLength returns the total length of all text (main + footnotes + headers + etc.).
func (wds *WordDocumentStream) GetTotalTextLength() uint32 {
	return wds.FIB.FibRgLw.CcpText +
		wds.FIB.FibRgLw.CcpFtn +
		wds.FIB.FibRgLw.CcpHdd +
		wds.FIB.FibRgLw.CcpAtn +
		wds.FIB.FibRgLw.CcpEdn +
		wds.FIB.FibRgLw.CcpTxbx +
		wds.FIB.FibRgLw.CcpHdrTxbx
}

// IsEncrypted returns true if the document is encrypted.
func (wds *WordDocumentStream) IsEncrypted() bool {
	return wds.FIB.IsEncrypted()
}

// IsObfuscated returns true if the document uses XOR obfuscation.
func (wds *WordDocumentStream) IsObfuscated() bool {
	return wds.FIB.IsObfuscated()
}

// Common stream errors
var (
	ErrStreamTooSmall = fmt.Errorf("stream data too small for requested range")
)
