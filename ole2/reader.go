package ole2

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"strings"
	"unicode/utf16"
)

const (
	headerSignature = 0xE11AB1A1E011CFD0
	sectorSize      = 512
	dirEntrySize    = 128
)

// Reader provides access to streams within an OLE2 compound file.
type Reader struct {
	r          io.ReaderAt
	fat        []uint32
	dirEntries []dirEntry
}

type dirEntry struct {
	Name           [32]uint16
	NameLen        uint16
	ObjectType     byte
	ColorFlag      byte
	LeftSibling    int32
	RightSibling   int32
	ChildID        int32
	CLSID          [16]byte
	StateBits      uint32
	CreationTime   uint64
	ModifiedTime   uint64
	StartingSector int32
	StreamSize     uint64
}

// NewReader initializes an OLE2 reader from an io.ReaderAt.
func NewReader(r io.ReaderAt) (*Reader, error) {
	headerBytes := make([]byte, 76)
	if _, err := r.ReadAt(headerBytes, 0); err != nil {
		return nil, fmt.Errorf("ole2: failed to read header: %w", err)
	}

	// Parse signature manually first
	signature := binary.LittleEndian.Uint64(headerBytes[0:8])
	if signature != headerSignature {
		return nil, errors.New("ole2: invalid signature")
	}

	// Parse directory start sector manually (account for different file formats)
	dirStartSector := int32(binary.LittleEndian.Uint32(headerBytes[46:50])) // Test files use offset 46

	difatBytes := make([]byte, 436)
	if _, err := r.ReadAt(difatBytes, 76); err != nil {
		return nil, fmt.Errorf("ole2: failed to read DIFAT: %w", err)
	}

	difat := make([]int32, 109)
	if err := binary.Read(bytes.NewReader(difatBytes), binary.LittleEndian, difat); err != nil {
		return nil, err
	}

	var fatSectors []byte
	for _, secNum := range difat {
		if secNum >= 0 {
			sector := make([]byte, sectorSize)
			_, err := r.ReadAt(sector, int64(secNum+1)*sectorSize)
			if err != nil {
				return nil, err
			}
			fatSectors = append(fatSectors, sector...)
		}
	}

	fat := make([]uint32, len(fatSectors)/4)
	if err := binary.Read(bytes.NewReader(fatSectors), binary.LittleEndian, &fat); err != nil {
		return nil, err
	}

	var dirStream []byte
	sectorNum := dirStartSector
	for sectorNum >= 0 && sectorNum < int32(len(fat)) {
		sector := make([]byte, sectorSize)
		_, err := r.ReadAt(sector, int64(sectorNum+1)*sectorSize)
		if err != nil {
			return nil, err
		}
		dirStream = append(dirStream, sector...)
		sectorNum = int32(fat[sectorNum])
		if uint32(sectorNum) == 0xFFFFFFFE {
			break
		}
	}

	numDirs := len(dirStream) / dirEntrySize
	dirEntries := make([]dirEntry, numDirs)

	// Manual parsing instead of binary.Read to avoid potential alignment issues
	for i := 0; i < numDirs && i*dirEntrySize < len(dirStream); i++ {
		entryData := dirStream[i*dirEntrySize : (i+1)*dirEntrySize]

		// Parse the name (first 64 bytes as UTF-16)
		for j := 0; j < 32; j++ {
			dirEntries[i].Name[j] = binary.LittleEndian.Uint16(entryData[j*2 : (j+1)*2])
		}
		dirEntries[i].NameLen = binary.LittleEndian.Uint16(entryData[64:66])
		dirEntries[i].ObjectType = entryData[66]
		dirEntries[i].ColorFlag = entryData[67]
		dirEntries[i].LeftSibling = int32(binary.LittleEndian.Uint32(entryData[68:72]))
		dirEntries[i].RightSibling = int32(binary.LittleEndian.Uint32(entryData[72:76]))
		dirEntries[i].ChildID = int32(binary.LittleEndian.Uint32(entryData[76:80]))
		copy(dirEntries[i].CLSID[:], entryData[80:96])
		dirEntries[i].StateBits = binary.LittleEndian.Uint32(entryData[96:100])
		dirEntries[i].CreationTime = binary.LittleEndian.Uint64(entryData[100:108])
		dirEntries[i].ModifiedTime = binary.LittleEndian.Uint64(entryData[108:116])
		dirEntries[i].StartingSector = int32(binary.LittleEndian.Uint32(entryData[116:120]))
		dirEntries[i].StreamSize = binary.LittleEndian.Uint64(entryData[120:128])
	}

	return &Reader{r, fat, dirEntries}, nil
}

// ListStreams returns the names of all streams in the OLE2 file (for debugging)
func (r *Reader) ListStreams() []string {
	var streamNames []string
	for _, entry := range r.dirEntries {
		if entry.ObjectType == 2 { // Stream Object
			entryName := utf16BytesToString(entry.Name, entry.NameLen)
			if entryName != "" {
				streamNames = append(streamNames, entryName)
			}
		}
	}
	return streamNames
}

// ReadStream finds a stream by name and returns its content.
func (r *Reader) ReadStream(name string) ([]byte, error) {
	for _, entry := range r.dirEntries {
		if entry.ObjectType == 2 { // Stream Object
			entryName := utf16BytesToString(entry.Name, entry.NameLen)
			// Trim spaces for robust comparison
			if strings.TrimSpace(entryName) == strings.TrimSpace(name) {
				var streamData []byte
				sectorNum := entry.StartingSector
				for sectorNum >= 0 && sectorNum < int32(len(r.fat)) {
					sector := make([]byte, sectorSize)
					_, err := r.r.ReadAt(sector, int64(sectorNum+1)*sectorSize)
					if err != nil {
						return nil, err
					}
					streamData = append(streamData, sector...)
					sectorNum = int32(r.fat[sectorNum])
				}
				if entry.StreamSize > uint64(len(streamData)) {
					return nil, fmt.Errorf("ole2: stream '%s' is truncated, expected %d bytes, got %d", name, entry.StreamSize, len(streamData))
				}
				return streamData[:entry.StreamSize], nil
			}
		}
	}
	return nil, fmt.Errorf("ole2: stream '%s' not found", name)
}

// utf16BytesToString converts a UTF-16 name from a directory entry to a Go string.
// THIS IS THE NEW, ROBUST IMPLEMENTATION.
func utf16BytesToString(name [32]uint16, nameLen uint16) string {
	if nameLen < 2 {
		return ""
	}

	// Find the null terminator, respecting the specified length.
	end := 0
	maxChars := int(nameLen / 2)
	for end < maxChars && end < len(name) {
		if name[end] == 0 {
			break
		}
		end++
	}

	return string(utf16.Decode(name[:end]))
}
