package tests

import (
    "bytes"
    "encoding/binary"
    "os"
    "testing"
    "unicode/utf16"

    "github.com/TalentFormula/msdoc/pkg/msdoc"
)

// createMockDocFile creates a complete, minimal .doc file in a byte buffer
// containing the text "Hello World".
func createMockDocFile(t *testing.T) []byte {
    var buf bytes.Buffer
    sectorSize := 512

    // --- Text content ---
    textContent := "Hello World"
    utf16Content := utf16.Encode([]rune(textContent))
    utf16Bytes := make([]byte, len(utf16Content)*2)
    for i, r := range utf16Content {
        binary.LittleEndian.PutUint16(utf16Bytes[i*2:], r)
    }

    // --- OLE2 Header and DIFAT (Sector -1) ---
    header := make([]byte, sectorSize)
    binary.LittleEndian.PutUint64(header[0:], 0xE011CFD0B1A1E11A) // Signature
    binary.LittleEndian.PutUint16(header[28:], 0x0009)            // 512-byte sectors
    binary.LittleEndian.PutUint32(header[44:], 1)                 // Directory stream starts at sector 1
    binary.LittleEndian.PutUint32(header[76:], 0)                 // First FAT sector is at sector 0
    for i := 80; i < sectorSize; i += 4 {
        binary.LittleEndian.PutUint32(header[i:], 0xFFFFFFFF) // End of DIFAT
    }
    buf.Write(header)

    // --- FAT (Sector 0) ---
    fat := make([]byte, sectorSize)
    binary.LittleEndian.PutUint32(fat[0:], 0xFFFFFFFD)  // FAT sector
    binary.LittleEndian.PutUint32(fat[4:], 0xFFFFFFFE)  // End of directory chain
    binary.LittleEndian.PutUint32(fat[8:], 4)           // WordDocument chain (sector 2 -> 4)
    binary.LittleEndian.PutUint32(fat[12:], 0xFFFFFFFE) // End of 0Table chain
    binary.LittleEndian.PutUint32(fat[16:], 0xFFFFFFFE) // End of WordDocument chain
    buf.Write(fat)

    // --- Directory Stream (Sector 1) ---
    dirStream := make([]byte, sectorSize)
    // Entry 0: Root Entry
    rootName := strToUtf16("Root Entry")
    for i, r := range rootName {
        binary.LittleEndian.PutUint16(dirStream[i*2:], r)
    }
    dirStream[66] = 5
    binary.LittleEndian.PutUint32(dirStream[76:], uint32(1)) // Child ID
    // Entry 1: WordDocument
    wdName := strToUtf16("WordDocument")
    for i, r := range wdName {
        binary.LittleEndian.PutUint16(dirStream[128+i*2:], r)
    }
    dirStream[128+66] = 2
    binary.LittleEndian.PutUint16(dirStream[128+64:], uint16(len(wdName)*2))
    binary.LittleEndian.PutUint32(dirStream[128+116:], uint32(2)) // Starts at sector 2
    binary.LittleEndian.PutUint64(dirStream[128+120:], uint64(sectorSize+len(utf16Bytes)))
    // Entry 2: 0Table
    tableName := strToUtf16("0Table")
    for i, r := range tableName {
        binary.LittleEndian.PutUint16(dirStream[256+i*2:], r)
    }
    dirStream[256+66] = 2
    binary.LittleEndian.PutUint16(dirStream[256+64:], uint16(len(tableName)*2))
    binary.LittleEndian.PutUint32(dirStream[256+116:], uint32(3)) // Starts at sector 3
    binary.LittleEndian.PutUint64(dirStream[256+120:], 512)
    buf.Write(dirStream)

    // --- WordDocument Stream (Sector 2 and 4) ---
    wordDocStream := make([]byte, sectorSize*2)
    // A minimal FIB
    binary.LittleEndian.PutUint16(wordDocStream[0:], 0xA5EC)          // wIdent
    binary.LittleEndian.PutUint16(wordDocStream[10:], 0x0000)         // fWhichTblStm = 0 (use 0Table)
    binary.LittleEndian.PutUint16(wordDocStream[32+2+28+2+88+2:], 93) // cbRgFcLcb
    // Place text content starting at offset 512
    copy(wordDocStream[512:], utf16Bytes)
    buf.Write(wordDocStream[:sectorSize]) // Sector 2

    // --- 0Table Stream (Sector 3) ---
    tableStream := make([]byte, sectorSize)
    // PlcPcd (Piece Table) pointing to the text
    tableStream[0] = 0x02 // Identifies a PlcPcd
    numPcds := 1
    // CPs: 0 and 11 (len("Hello World"))
    binary.LittleEndian.PutUint32(tableStream[1:], 0)
    binary.LittleEndian.PutUint32(tableStream[5:], uint32(len(textContent)))
    // PCD: describes the piece
    pcdOffset := 1 + 4*(numPcds+1)
    fc := uint32(512*2) | 0x40000000 // Offset 512, Unicode flag set
    binary.LittleEndian.PutUint32(tableStream[pcdOffset+2:], fc)
    buf.Write(tableStream)

    // Write the second part of the WordDocument stream
    buf.Write(wordDocStream[sectorSize:]) // Sector 4

    return buf.Bytes()
}

func TestReaderAPI(t *testing.T) {
    // 1. Create a temporary file with the mock doc content
    mockData := createMockDocFile(t)
    tmpfile, err := os.CreateTemp("", "sample-*.doc")
    if err != nil {
        t.Fatal(err)
    }
    defer os.Remove(tmpfile.Name())

    if _, err := tmpfile.Write(mockData); err != nil {
        t.Fatal(err)
    }
    if err := tmpfile.Close(); err != nil {
        t.Fatal(err)
    }

    // 2. Use the library to open and parse it
    doc, err := msdoc.Open(tmpfile.Name())
    if err != nil {
        t.Fatalf("msdoc.Open failed: %v", err)
    }
    defer doc.Close()

    // 3. Extract text and verify
    text, err := doc.Text()
    if err != nil {
        t.Fatalf("doc.Text() failed: %v", err)
    }

    expectedText := "Hello World"
    if text != expectedText {
        t.Errorf("Expected text '%s', got '%s'", expectedText, text)
    }
}
