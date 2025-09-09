package tests

import (
	"encoding/binary"
	"testing"

	"github.com/TalentFormula/msdoc/structures"
)

func TestPLCParsing(t *testing.T) {
	// Create a mock PLC with 3 data elements (4 CPs, 3 data elements of 8 bytes each)
	plcData := make([]byte, 4*4+3*8) // 4 CPs + 3 data elements

	// Write CPs
	binary.LittleEndian.PutUint32(plcData[0:], 0)    // CP 0
	binary.LittleEndian.PutUint32(plcData[4:], 100)  // CP 100
	binary.LittleEndian.PutUint32(plcData[8:], 200)  // CP 200
	binary.LittleEndian.PutUint32(plcData[12:], 300) // CP 300

	// Write data elements (8 bytes each)
	for i := 0; i < 3; i++ {
		offset := 16 + i*8
		binary.LittleEndian.PutUint64(plcData[offset:], uint64(i+1)*1000)
	}

	// Parse the PLC
	plc, err := structures.ParsePLC(plcData, 8)
	if err != nil {
		t.Fatalf("ParsePLC failed: %v", err)
	}

	// Validate structure
	if err := plc.Validate(); err != nil {
		t.Errorf("PLC validation failed: %v", err)
	}

	// Check counts
	if plc.Count() != 3 {
		t.Errorf("Expected 3 data elements, got %d", plc.Count())
	}

	if len(plc.CPs) != 4 {
		t.Errorf("Expected 4 CPs, got %d", len(plc.CPs))
	}

	// Check CP values
	expectedCPs := []structures.CP{0, 100, 200, 300}
	for i, expected := range expectedCPs {
		if plc.CPs[i] != expected {
			t.Errorf("CP[%d]: expected %d, got %d", i, expected, plc.CPs[i])
		}
	}

	// Check ranges
	for i := 0; i < 3; i++ {
		start, end, err := plc.GetRange(i)
		if err != nil {
			t.Errorf("GetRange(%d) failed: %v", i, err)
		}
		expectedStart := expectedCPs[i]
		expectedEnd := expectedCPs[i+1]
		if start != expectedStart || end != expectedEnd {
			t.Errorf("Range[%d]: expected (%d, %d), got (%d, %d)",
				i, expectedStart, expectedEnd, start, end)
		}
	}
}

func TestPLCInvalidData(t *testing.T) {
	// Test with invalid data size
	invalidData := []byte{1, 2, 3} // Too small
	_, err := structures.ParsePLC(invalidData, 8)
	if err == nil {
		t.Error("Expected error for invalid data size")
	}

	// Test with mismatched size
	mismatchedData := make([]byte, 17) // Should be multiple of (dataSize + 4)
	_, err = structures.ParsePLC(mismatchedData, 8)
	if err == nil {
		t.Error("Expected error for mismatched data size")
	}
}

func TestPCDParsing(t *testing.T) {
	// Create a mock PCD (8 bytes)
	pcdData := make([]byte, 8)

	// Set flags (first 2 bytes)
	flags := uint16(0x0003) // fNoEncryption | fComplex
	binary.LittleEndian.PutUint16(pcdData[0:], flags)

	// Set FC (next 4 bytes) with Unicode flag
	fc := uint32(0x40001000) // Unicode text at position 0x1000
	binary.LittleEndian.PutUint32(pcdData[2:], fc)

	// Parse PCD
	pcd, err := structures.ParsePCD(pcdData)
	if err != nil {
		t.Fatalf("ParsePCD failed: %v", err)
	}

	// Check flags
	if !pcd.FNoEncryption {
		t.Error("Expected FNoEncryption to be true")
	}
	if !pcd.FComplex {
		t.Error("Expected FComplex to be true")
	}
	if !pcd.IsUnicode {
		t.Error("Expected IsUnicode to be true")
	}

	// Check FC
	expectedFC := uint32(0x1000)
	if pcd.FC != expectedFC {
		t.Errorf("Expected FC %d, got %d", expectedFC, pcd.FC)
	}

	// Check actual FC calculation
	actualFC := pcd.GetActualFC()
	expectedActualFC := uint32(0x1000 / 2) // Unicode FC is divided by 2
	if actualFC != expectedActualFC {
		t.Errorf("Expected actual FC %d, got %d", expectedActualFC, actualFC)
	}
}

func TestPlcPcdParsing(t *testing.T) {
	// Create a mock PlcPcd with 2 pieces
	numPieces := 2
	plcSize := (numPieces+1)*4 + numPieces*8 // 3 CPs + 2 PCDs
	plcData := make([]byte, plcSize)

	// Write CPs
	binary.LittleEndian.PutUint32(plcData[0:], 0)   // CP 0
	binary.LittleEndian.PutUint32(plcData[4:], 100) // CP 100
	binary.LittleEndian.PutUint32(plcData[8:], 200) // CP 200

	// Write PCDs
	// PCD 1: ANSI text
	offset := 12
	binary.LittleEndian.PutUint16(plcData[offset:], 0x0001)   // fNoEncryption
	binary.LittleEndian.PutUint32(plcData[offset+2:], 0x2000) // FC (ANSI)

	// PCD 2: Unicode text
	offset = 20
	binary.LittleEndian.PutUint16(plcData[offset:], 0x0000)       // No flags
	binary.LittleEndian.PutUint32(plcData[offset+2:], 0x40003000) // FC (Unicode)

	// Parse PlcPcd
	plcPcd, err := structures.ParsePlcPcd(plcData)
	if err != nil {
		t.Fatalf("ParsePlcPcd failed: %v", err)
	}

	// Check count
	if plcPcd.Count() != 2 {
		t.Errorf("Expected 2 pieces, got %d", plcPcd.Count())
	}

	// Check first piece (ANSI)
	start, end, pcd1, err := plcPcd.GetTextRange(0)
	if err != nil {
		t.Fatalf("GetTextRange(0) failed: %v", err)
	}
	if start != 0 || end != 100 {
		t.Errorf("Piece 0: expected range (0, 100), got (%d, %d)", start, end)
	}
	if pcd1.IsUnicode {
		t.Error("Piece 0 should be ANSI")
	}
	if pcd1.FC != 0x2000 {
		t.Errorf("Piece 0: expected FC 0x2000, got 0x%X", pcd1.FC)
	}

	// Check second piece (Unicode)
	start, end, pcd2, err := plcPcd.GetTextRange(1)
	if err != nil {
		t.Fatalf("GetTextRange(1) failed: %v", err)
	}
	if start != 100 || end != 200 {
		t.Errorf("Piece 1: expected range (100, 200), got (%d, %d)", start, end)
	}
	if !pcd2.IsUnicode {
		t.Error("Piece 1 should be Unicode")
	}
	if pcd2.FC != 0x3000 {
		t.Errorf("Piece 1: expected FC 0x3000, got 0x%X", pcd2.FC)
	}
}
