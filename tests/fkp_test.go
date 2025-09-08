package tests

import (
	"encoding/binary"
	"testing"

	"github.com/TalentFormula/msdoc/structures"
)

func TestFKPBasicParsing(t *testing.T) {
	// Create a mock 512-byte FKP
	fkpData := make([]byte, 512)
	
	// Set entry count (last byte)
	entryCount := byte(2)
	fkpData[511] = entryCount

	// Parse as unknown type
	fkp, err := structures.ParseFKP(fkpData, structures.FKPTypeUnknown)
	if err != nil {
		t.Fatalf("ParseFKP failed: %v", err)
	}

	if fkp.EntryCount != int(entryCount) {
		t.Errorf("Expected entry count %d, got %d", entryCount, fkp.EntryCount)
	}

	if fkp.Type != structures.FKPTypeUnknown {
		t.Errorf("Expected type %d, got %d", structures.FKPTypeUnknown, fkp.Type)
	}

	if len(fkp.Data) != 512 {
		t.Errorf("Expected data size 512, got %d", len(fkp.Data))
	}
}

func TestCHPXFKPParsing(t *testing.T) {
	// Create a mock CHPX FKP with 2 entries
	fkpData := make([]byte, 512)
	
	// Entry 1: FC=100, offset=200
	binary.LittleEndian.PutUint32(fkpData[0:], 100)
	fkpData[4] = 200
	
	// Entry 2: FC=200, offset=210  
	binary.LittleEndian.PutUint32(fkpData[5:], 200)
	fkpData[9] = 210

	// Add formatting data at offset 200 (length=8)
	fkpData[200] = 8 // Length byte
	for i := 0; i < 8; i++ {
		fkpData[201+i] = byte(i + 1) // Test data
	}

	// Add formatting data at offset 210 (length=4)
	fkpData[210] = 4 // Length byte
	for i := 0; i < 4; i++ {
		fkpData[211+i] = byte(i + 10) // Test data
	}

	// Set entry count
	fkpData[511] = 2

	// Parse as CHPX FKP
	fkp, err := structures.ParseFKP(fkpData, structures.FKPTypeCHP)
	if err != nil {
		t.Fatalf("ParseFKP failed: %v", err)
	}

	if len(fkp.Entries) != 2 {
		t.Errorf("Expected 2 entries, got %d", len(fkp.Entries))
	}

	// Check first entry
	entry1, err := fkp.GetEntryAt(0)
	if err != nil {
		t.Fatalf("GetEntryAt(0) failed: %v", err)
	}
	if entry1.FC != 100 {
		t.Errorf("Entry 0: expected FC 100, got %d", entry1.FC)
	}
	if entry1.Offset != 200 {
		t.Errorf("Entry 0: expected offset 200, got %d", entry1.Offset)
	}
	if len(entry1.Data) != 8 {
		t.Errorf("Entry 0: expected data length 8, got %d", len(entry1.Data))
	}
	for i, expected := range []byte{1, 2, 3, 4, 5, 6, 7, 8} {
		if entry1.Data[i] != expected {
			t.Errorf("Entry 0 data[%d]: expected %d, got %d", i, expected, entry1.Data[i])
		}
	}

	// Check second entry
	entry2, err := fkp.GetEntryAt(1)
	if err != nil {
		t.Fatalf("GetEntryAt(1) failed: %v", err)
	}
	if entry2.FC != 200 {
		t.Errorf("Entry 1: expected FC 200, got %d", entry2.FC)
	}
	if len(entry2.Data) != 4 {
		t.Errorf("Entry 1: expected data length 4, got %d", len(entry2.Data))
	}
}

func TestPAPXFKPParsing(t *testing.T) {
	// Create a mock PAPX FKP with 1 entry
	fkpData := make([]byte, 512)
	
	// Entry 1: FC=300, offset=220
	binary.LittleEndian.PutUint32(fkpData[0:], 300)
	binary.LittleEndian.PutUint16(fkpData[4:], 220)

	// Add formatting data at offset 220 (length=3 words = 6 bytes)
	fkpData[220] = 3 // Length in words
	for i := 0; i < 6; i++ {
		fkpData[221+i] = byte(i + 20) // Test data
	}

	// Set entry count
	fkpData[511] = 1

	// Parse as PAPX FKP
	fkp, err := structures.ParseFKP(fkpData, structures.FKPTypePAP)
	if err != nil {
		t.Fatalf("ParseFKP failed: %v", err)
	}

	if len(fkp.Entries) != 1 {
		t.Errorf("Expected 1 entry, got %d", len(fkp.Entries))
	}

	// Check entry
	entry, err := fkp.GetEntryAt(0)
	if err != nil {
		t.Fatalf("GetEntryAt(0) failed: %v", err)
	}
	if entry.FC != 300 {
		t.Errorf("Expected FC 300, got %d", entry.FC)
	}
	if entry.Offset != 220 {
		t.Errorf("Expected offset 220, got %d", entry.Offset)
	}
	if len(entry.Data) != 6 {
		t.Errorf("Expected data length 6, got %d", len(entry.Data))
	}
	for i, expected := range []byte{20, 21, 22, 23, 24, 25} {
		if entry.Data[i] != expected {
			t.Errorf("Entry data[%d]: expected %d, got %d", i, expected, entry.Data[i])
		}
	}
}

func TestFKPFindEntryForFC(t *testing.T) {
	// Create a mock CHPX FKP with multiple entries
	fkpData := make([]byte, 512)
	
	// Entry 1: FC=100
	binary.LittleEndian.PutUint32(fkpData[0:], 100)
	fkpData[4] = 0 // No formatting data
	
	// Entry 2: FC=200
	binary.LittleEndian.PutUint32(fkpData[5:], 200)
	fkpData[9] = 0 // No formatting data
	
	// Entry 3: FC=300
	binary.LittleEndian.PutUint32(fkpData[10:], 300)
	fkpData[14] = 0 // No formatting data

	// Set entry count
	fkpData[511] = 3

	// Parse FKP
	fkp, err := structures.ParseFKP(fkpData, structures.FKPTypeCHP)
	if err != nil {
		t.Fatalf("ParseFKP failed: %v", err)
	}

	// Test finding entries for various FC values
	testCases := []struct {
		fc       uint32
		expected uint32 // Expected FC of found entry
		found    bool
	}{
		{50, 0, false},    // Before first entry
		{100, 100, true},  // Exact match with first
		{150, 100, true},  // Between first and second
		{200, 200, true},  // Exact match with second
		{250, 200, true},  // Between second and third
		{300, 300, true},  // Exact match with third
		{400, 300, true},  // After last entry
	}

	for _, tc := range testCases {
		entry := fkp.FindEntryForFC(tc.fc)
		if tc.found {
			if entry == nil {
				t.Errorf("FC %d: expected to find entry, got nil", tc.fc)
			} else if entry.FC != tc.expected {
				t.Errorf("FC %d: expected entry FC %d, got %d", tc.fc, tc.expected, entry.FC)
			}
		} else {
			if entry != nil {
				t.Errorf("FC %d: expected no entry, got FC %d", tc.fc, entry.FC)
			}
		}
	}
}

func TestFKPInvalidData(t *testing.T) {
	// Test with wrong size
	invalidData := make([]byte, 256) // Should be 512
	_, err := structures.ParseFKP(invalidData, structures.FKPTypeCHP)
	if err == nil {
		t.Error("Expected error for invalid data size")
	}

	// Test with too many entries for available space
	fkpData := make([]byte, 512)
	fkpData[511] = 200 // 200 * 5 = 1000 bytes > 511 available
	_, err = structures.ParseFKP(fkpData, structures.FKPTypeCHP)
	if err == nil {
		t.Error("Expected error for too many entries")
	}
}