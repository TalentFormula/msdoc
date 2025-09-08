package structures

import (
	"encoding/binary"
	"fmt"
)

// PLC (Plex) is a common structure in .doc files. It is an array of
// Character Positions (CPs) followed by an array of data elements.
// The number of CPs is always one more than the number of data elements.
type PLC struct {
	CPs      []CP
	Data     [][]byte // Generic representation of data elements
	DataSize int      // Size of each data element in bytes
}

// ParsePLC parses a PLC structure from raw bytes.
// dataSize specifies the size of each data element in bytes.
func ParsePLC(data []byte, dataSize int) (*PLC, error) {
	if len(data) < 4 {
		return nil, fmt.Errorf("plc: data too short, need at least 4 bytes")
	}

	if dataSize <= 0 {
		return nil, fmt.Errorf("plc: invalid data size %d", dataSize)
	}

	// Calculate number of data elements
	// Formula: n = (cbPlc - 4) / (dataSize + 4)
	// where cbPlc is the total PLC size, dataSize is size of each data element,
	// and 4 is the size of each CP (32-bit integer)
	if (len(data)-4)%(dataSize+4) != 0 {
		return nil, fmt.Errorf("plc: invalid PLC size %d for data element size %d", len(data), dataSize)
	}

	numDataElements := (len(data) - 4) / (dataSize + 4)
	numCPs := numDataElements + 1

	// Parse CPs
	cps := make([]CP, numCPs)
	for i := 0; i < numCPs; i++ {
		offset := i * 4
		if offset+4 > len(data) {
			return nil, fmt.Errorf("plc: not enough data for CP %d", i)
		}
		cps[i] = CP(binary.LittleEndian.Uint32(data[offset : offset+4]))
	}

	// Parse data elements
	dataElements := make([][]byte, numDataElements)
	dataOffset := numCPs * 4
	for i := 0; i < numDataElements; i++ {
		offset := dataOffset + (i * dataSize)
		if offset+dataSize > len(data) {
			return nil, fmt.Errorf("plc: not enough data for element %d", i)
		}
		dataElements[i] = make([]byte, dataSize)
		copy(dataElements[i], data[offset:offset+dataSize])
	}

	return &PLC{
		CPs:      cps,
		Data:     dataElements,
		DataSize: dataSize,
	}, nil
}

// Count returns the number of data elements in the PLC.
func (plc *PLC) Count() int {
	return len(plc.Data)
}

// GetRange returns the character range for the given data element index.
func (plc *PLC) GetRange(index int) (start, end CP, err error) {
	if index < 0 || index >= len(plc.Data) {
		return 0, 0, fmt.Errorf("plc: invalid index %d", index)
	}
	return plc.CPs[index], plc.CPs[index+1], nil
}

// GetDataAt returns the data element at the given index.
func (plc *PLC) GetDataAt(index int) ([]byte, error) {
	if index < 0 || index >= len(plc.Data) {
		return nil, fmt.Errorf("plc: invalid index %d", index)
	}
	return plc.Data[index], nil
}

// Validate performs basic validation on the PLC structure.
func (plc *PLC) Validate() error {
	if len(plc.CPs) != len(plc.Data)+1 {
		return fmt.Errorf("plc: invalid structure, CPs count (%d) != Data count (%d) + 1", len(plc.CPs), len(plc.Data))
	}

	// Check that CPs are in ascending order
	for i := 1; i < len(plc.CPs); i++ {
		if plc.CPs[i] < plc.CPs[i-1] {
			return fmt.Errorf("plc: CPs not in ascending order at index %d", i)
		}
	}

	// Check that all data elements have the same size
	for i, data := range plc.Data {
		if len(data) != plc.DataSize {
			return fmt.Errorf("plc: data element %d has size %d, expected %d", i, len(data), plc.DataSize)
		}
	}

	return nil
}
