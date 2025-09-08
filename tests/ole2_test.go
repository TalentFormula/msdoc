package tests

import (
	"bytes"
	"encoding/binary"
	"testing"
	"unicode/utf16"

	"github.com/TalentFormula/msdoc/ole2"
)

// Helper to create a UTF-16 representation for directory entries
func strToUtf16(s string) []uint16 {
	return utf16.Encode([]rune(s + "\x00"))
}

func TestOLE2Reader(t *testing.T) {
	// --- Create a mock OLE2 file in memory ---
	var buf bytes.Buffer
	sectorSize := 512

	// 1. OLE2 Header (76 bytes)
	header := make([]byte, 76)
	binary.LittleEndian.PutUint64(header[0:], 0xE011CFD0B1A1E11A) // Signature
	binary.LittleEndian.PutUint16(header[28:], 0x0009)            // Sector Shift (512 bytes)
	binary.LittleEndian.PutUint32(header[44:], 1)                 // Directory Start Sector
	buf.Write(header)

	// 2. DIFAT (rest of the first sector)
	difat := make([]byte, sectorSize-76)
	for i := range difat {
		difat[i] = 0xFF // Unused
	}
	binary.LittleEndian.PutUint32(difat[0:], 0) // FAT is in sector 0
	buf.Write(difat)

	// 3. FAT Sector (Sector 0)
	fat := make([]byte, sectorSize)
	binary.LittleEndian.PutUint32(fat[0:], 0xFFFFFFFD) // FAT sector marker
	binary.LittleEndian.PutUint32(fat[4:], 0xFFFFFFFE) // Directory chain end
	binary.LittleEndian.PutUint32(fat[8:], 0xFFFFFFFE) // Stream chain end
	buf.Write(fat)

	// 4. Directory Sector (Sector 1)
	dirSector := make([]byte, sectorSize)
	// Root Entry (Entry 0)
	rootName := strToUtf16("Root Entry")
	for i, r := range rootName {
		binary.LittleEndian.PutUint16(dirSector[i*2:], r)
	}
	binary.LittleEndian.PutUint16(dirSector[64:], uint16(len(rootName)*2))
	dirSector[66] = 5                                        // Object Type: Root
	binary.LittleEndian.PutUint32(dirSector[76:], uint32(1)) // Child ID: 1 (our stream)

	// Stream Entry (Entry 1)
	streamName := strToUtf16("MyStream")
	for i, r := range streamName {
		binary.LittleEndian.PutUint16(dirSector[128+i*2:], r)
	}
	binary.LittleEndian.PutUint16(dirSector[128+64:], uint16(len(streamName)*2))
	dirSector[128+66] = 2                                         // Object Type: Stream
	binary.LittleEndian.PutUint32(dirSector[128+116:], uint32(2)) // Starting Sector: 2
	binary.LittleEndian.PutUint64(dirSector[128+120:], 12)        // Stream Size: 12 bytes
	buf.Write(dirSector)

	// 5. Stream Data Sector (Sector 2)
	streamData := []byte("Hello OLE2!")
	streamSector := make([]byte, sectorSize)
	copy(streamSector, streamData)
	buf.Write(streamSector)

	// --- Run the test ---
	oleReader, err := ole2.NewReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("NewReader failed: %v", err)
	}

	data, err := oleReader.ReadStream("MyStream")
	if err != nil {
		t.Fatalf("ReadStream failed: %v", err)
	}

	if len(data) != 12 {
		t.Errorf("Expected stream size 12, got %d", len(data))
	}

	expected := "Hello OLE2!"
	if string(data[:11]) != expected {
		t.Errorf("Expected stream content '%s', got '%s'", expected, string(data))
	}
}
