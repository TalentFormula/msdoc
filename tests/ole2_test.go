package tests

import (
	"bytes"
	"encoding/binary"
	"os"
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
	binary.LittleEndian.PutUint64(header[0:], 0xE11AB1A1E011CFD0) // Signature
	binary.LittleEndian.PutUint16(header[28:], 0x0009)            // Sector Shift (512 bytes)
	binary.LittleEndian.PutUint32(header[48:], 1)                 // Directory Start Sector (correct offset per OLE2 spec)
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

	// Test ListStreams functionality
	streams := oleReader.ListStreams()
	if len(streams) != 1 {
		t.Errorf("Expected 1 stream, got %d", len(streams))
	}
	if len(streams) > 0 && streams[0] != "MyStream" {
		t.Errorf("Expected stream 'MyStream', got '%s'", streams[0])
	}

	// Test ReadStream functionality
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

func TestOLE2RealWordDocs(t *testing.T) {
	// Test with sample-1.doc
	file1, err := os.Open("testdata/sample-1.doc")
	if err != nil {
		t.Fatalf("Failed to open sample-1.doc: %v", err)
	}
	defer file1.Close()

	reader1, err := ole2.NewReader(file1)
	if err != nil {
		t.Fatalf("Failed to create OLE2 reader for sample-1.doc: %v", err)
	}

	streams1 := reader1.ListStreams()
	if len(streams1) == 0 {
		t.Errorf("Expected streams in sample-1.doc, but got none")
	}

	// Check for expected Word document streams
	expectedStreams := []string{"WordDocument", "1Table"}
	for _, expectedStream := range expectedStreams {
		found := false
		for _, stream := range streams1 {
			if stream == expectedStream {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("Expected stream '%s' not found in sample-1.doc. Found streams: %v", expectedStream, streams1)
		}
	}

	// Test with sample-2.doc
	file2, err := os.Open("testdata/sample-2.doc")
	if err != nil {
		t.Fatalf("Failed to open sample-2.doc: %v", err)
	}
	defer file2.Close()

	reader2, err := ole2.NewReader(file2)
	if err != nil {
		t.Fatalf("Failed to create OLE2 reader for sample-2.doc: %v", err)
	}

	streams2 := reader2.ListStreams()
	if len(streams2) == 0 {
		t.Errorf("Expected streams in sample-2.doc, but got none")
	}

	// Check for expected Word document streams
	for _, expectedStream := range expectedStreams {
		found := false
		for _, stream := range streams2 {
			if stream == expectedStream {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("Expected stream '%s' not found in sample-2.doc. Found streams: %v", expectedStream, streams2)
		}
	}
}
