package ole2

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
)

// Writer provides functionality for creating OLE2 compound documents.
type Writer struct {
	streams map[string][]byte
	header  CompoundFileHeader
}

// CompoundFileHeader represents the OLE2 compound file header.
type CompoundFileHeader struct {
	Signature            [8]byte     // OLE2 signature
	MinorVersion         uint16      // Minor version
	MajorVersion         uint16      // Major version
	ByteOrder            uint16      // Byte order identifier
	SectorSize           uint16      // Sector size (power of 2)
	MiniSectorSize       uint16      // Mini sector size (power of 2)
	Reserved1            uint16      // Reserved field
	Reserved2            uint16      // Reserved field
	NumDirectorySectors  uint32      // Number of directory sectors
	NumFATSectors        uint32      // Number of FAT sectors
	DirectoryFirstSector uint32      // First directory sector
	TransactionSignature uint32      // Transaction signature
	MiniStreamCutoff     uint32      // Mini stream cutoff
	MiniFATFirstSector   uint32      // First mini FAT sector
	NumMiniFATSectors    uint32      // Number of mini FAT sectors
	DIFATFirstSector     uint32      // First DIFAT sector
	NumDIFATSectors      uint32      // Number of DIFAT sectors
	DIFAT                [109]uint32 // First 109 DIFAT entries
}

// NewWriter creates a new OLE2 writer.
func NewWriter() *Writer {
	writer := &Writer{
		streams: make(map[string][]byte),
	}

	// Initialize header with standard values
	copy(writer.header.Signature[:], []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	writer.header.MinorVersion = 0x003E
	writer.header.MajorVersion = 0x003E
	writer.header.ByteOrder = 0xFFFE
	writer.header.SectorSize = 9     // 512 bytes (2^9)
	writer.header.MiniSectorSize = 6 // 64 bytes (2^6)
	writer.header.MiniStreamCutoff = 4096

	return writer
}

// AddStream adds a stream to the compound document.
func (w *Writer) AddStream(name string, data []byte) {
	w.streams[name] = data
}

// WriteTo writes the complete compound document to the writer.
func (w *Writer) WriteTo(writer io.Writer) error {
	// Calculate sector size
	sectorSize := 1 << w.header.SectorSize // 512 bytes

	// Build directory entries
	dirEntries, err := w.buildDirectoryEntries()
	if err != nil {
		return fmt.Errorf("failed to build directory entries: %w", err)
	}

	// Calculate sectors needed
	totalDataSize := 0
	for _, data := range w.streams {
		totalDataSize += len(data)
	}

	numDataSectors := (totalDataSize + sectorSize - 1) / sectorSize
	numDirSectors := (len(dirEntries)*128 + sectorSize - 1) / sectorSize
	numFATSectors := ((numDataSectors+numDirSectors+1)*4 + sectorSize - 1) / sectorSize

	// Update header
	w.header.NumDirectorySectors = uint32(numDirSectors)
	w.header.NumFATSectors = uint32(numFATSectors)
	w.header.DirectoryFirstSector = uint32(numDataSectors)

	// Write header
	if err := binary.Write(writer, binary.LittleEndian, &w.header); err != nil {
		return fmt.Errorf("failed to write header: %w", err)
	}

	// Write data sectors
	currentSector := uint32(0)
	sectorMap := make(map[string]uint32)

	for name, data := range w.streams {
		sectorMap[name] = currentSector

		// Write data, padded to sector boundaries
		if _, err := writer.Write(data); err != nil {
			return fmt.Errorf("failed to write stream %s: %w", name, err)
		}

		// Pad to sector boundary
		padding := sectorSize - (len(data) % sectorSize)
		if padding < sectorSize {
			if _, err := writer.Write(make([]byte, padding)); err != nil {
				return fmt.Errorf("failed to write padding: %w", err)
			}
		}

		sectorsUsed := (len(data) + sectorSize - 1) / sectorSize
		currentSector += uint32(sectorsUsed)
	}

	// Write directory sectors
	dirData := w.buildDirectoryData(dirEntries, sectorMap)
	if _, err := writer.Write(dirData); err != nil {
		return fmt.Errorf("failed to write directory: %w", err)
	}

	// Pad directory to sector boundary
	dirPadding := (numDirSectors * sectorSize) - len(dirData)
	if dirPadding > 0 {
		if _, err := writer.Write(make([]byte, dirPadding)); err != nil {
			return fmt.Errorf("failed to write directory padding: %w", err)
		}
	}

	// Write FAT sectors
	fatData := w.buildFATData(numDataSectors, numDirSectors)
	if _, err := writer.Write(fatData); err != nil {
		return fmt.Errorf("failed to write FAT: %w", err)
	}

	return nil
}

// buildDirectoryEntries creates directory entries for all streams.
func (w *Writer) buildDirectoryEntries() ([]DirectoryEntry, error) {
	entries := make([]DirectoryEntry, 0)

	// Root entry
	rootEntry := DirectoryEntry{
		Name:         [64]uint16{0},
		NameLength:   10, // "Root Entry"
		Type:         5,  // Root storage
		NodeColor:    1,  // Red
		LeftSibling:  0xFFFFFFFF,
		RightSibling: 0xFFFFFFFF,
		Child:        1, // First stream
		StartSector:  0,
		Size:         0,
	}
	copy(rootEntry.Name[:], utf16Encode("Root Entry"))
	entries = append(entries, rootEntry)

	// Stream entries
	streamIndex := uint32(1)
	for name := range w.streams {
		entry := DirectoryEntry{
			Name:         [64]uint16{0},
			NameLength:   uint16((len(name) + 1) * 2),
			Type:         2, // Stream
			NodeColor:    0, // Black
			LeftSibling:  0xFFFFFFFF,
			RightSibling: 0xFFFFFFFF,
			Child:        0xFFFFFFFF,
		}
		copy(entry.Name[:], utf16Encode(name))
		entries = append(entries, entry)
		streamIndex++
	}

	return entries, nil
}

// DirectoryEntry represents an OLE2 directory entry.
type DirectoryEntry struct {
	Name         [64]uint16 // UTF-16 encoded name
	NameLength   uint16     // Length of name in bytes
	Type         uint8      // Entry type
	NodeColor    uint8      // Red-black tree node color
	LeftSibling  uint32     // Left sibling directory entry
	RightSibling uint32     // Right sibling directory entry
	Child        uint32     // Child directory entry
	CLSID        [16]byte   // Class ID
	StateBits    uint32     // State bits
	Created      uint64     // Creation time
	Modified     uint64     // Modified time
	StartSector  uint32     // Starting sector
	Size         uint64     // Size in bytes
}

// buildDirectoryData creates the directory data with sector mapping.
func (w *Writer) buildDirectoryData(entries []DirectoryEntry, sectorMap map[string]uint32) []byte {
	var buffer bytes.Buffer

	// Write root entry
	binary.Write(&buffer, binary.LittleEndian, &entries[0])

	// Write stream entries with proper sector assignments
	streamIndex := 1
	for name, data := range w.streams {
		if streamIndex < len(entries) {
			entry := entries[streamIndex]
			entry.StartSector = sectorMap[name]
			entry.Size = uint64(len(data))
			binary.Write(&buffer, binary.LittleEndian, &entry)
			streamIndex++
		}
	}

	return buffer.Bytes()
}

// buildFATData creates the File Allocation Table.
func (w *Writer) buildFATData(numDataSectors, numDirSectors int) []byte {
	var buffer bytes.Buffer

	// Mark data sectors as used
	for i := 0; i < numDataSectors-1; i++ {
		binary.Write(&buffer, binary.LittleEndian, uint32(i+1))
	}
	if numDataSectors > 0 {
		binary.Write(&buffer, binary.LittleEndian, uint32(0xFFFFFFFE)) // End of chain
	}

	// Mark directory sectors
	dirStart := uint32(numDataSectors)
	for i := 0; i < numDirSectors-1; i++ {
		binary.Write(&buffer, binary.LittleEndian, dirStart+uint32(i)+1)
	}
	if numDirSectors > 0 {
		binary.Write(&buffer, binary.LittleEndian, uint32(0xFFFFFFFE)) // End of chain
	}

	// Mark FAT sector as special
	binary.Write(&buffer, binary.LittleEndian, uint32(0xFFFFFFFD)) // FAT sector

	return buffer.Bytes()
}

// utf16Encode converts a string to UTF-16LE encoding.
func utf16Encode(s string) []uint16 {
	runes := []rune(s)
	result := make([]uint16, len(runes))
	for i, r := range runes {
		result[i] = uint16(r)
	}
	return result
}
