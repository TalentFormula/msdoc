package streams

import (
	"fmt"
)

// DataStream represents the optional Data stream containing additional document data.
type DataStream struct {
	Data []byte
}

// NewDataStream creates a new Data stream processor.
func NewDataStream(data []byte) *DataStream {
	return &DataStream{
		Data: data,
	}
}

// GetData extracts data from the specified range within the Data stream.
func (ds *DataStream) GetData(offset, length uint32) ([]byte, error) {
	if length == 0 {
		return nil, nil
	}

	if offset+length > uint32(len(ds.Data)) {
		return nil, fmt.Errorf("data: requested range out of bounds")
	}

	result := make([]byte, length)
	copy(result, ds.Data[offset:offset+length])
	return result, nil
}

// Size returns the total size of the Data stream.
func (ds *DataStream) Size() uint32 {
	return uint32(len(ds.Data))
}

// IsEmpty returns true if the Data stream has no content.
func (ds *DataStream) IsEmpty() bool {
	return len(ds.Data) == 0
}

// This file would contain logic for parsing the optional Data stream.
