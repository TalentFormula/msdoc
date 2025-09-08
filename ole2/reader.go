package ole2

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"unicode/utf16"
)

const (
	headerSignature = 0xE011CFD0B1A1E11A
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
	var header struct {
		Signature         uint64
		_                 [16]byte
		MinorVersion      uint16
		MajorVersion      uint16
		ByteOrder         uint16
		SectorShift       uint16
		_                 [10]byte
		NumFATSectors     uint32
		DirStartSector    int32
		_                 [4]byte
		MiniStreamCutoff  uint32
		MiniFATStart      int32
		NumMiniFATSectors uint32
		DIFATStart        int32
		NumDIFATSectors   uint32
	}

	headerBytes := make([]byte, 76)
	if _, err := r.ReadAt(headerBytes, 0); err != nil {
		return nil, fmt.Errorf("ole2: failed to read header: %w", err)
	}

	if err := binary.Read(bytes.NewReader(headerBytes), binary.LittleEndian, &header); err != nil {
		return nil, err
	}

	if header.Signature != headerSignature {
		return nil, errors.New("ole2: invalid signature")
	}

	difatBytes := make([]byte, 436)
	if _, err := r.ReadAt(difatBytes, 76); err != nil {
		return nil, fmt.Errorf("ole2: failed to read DIFAT: %w", err)
	}

	difat := make([]int32, 109)
	if err := binary.Read(bytes.NewReader(difatBytes), binary.LittleEndian, &difat); err != nil {
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
	sectorNum := header.DirStartSector
	for sectorNum >= 0 && sectorNum < int32(len(fat)) {
		sector := make([]byte, sectorSize)
		_, err := r.ReadAt(sector, int64(sectorNum+1)*sectorSize)
		if err != nil {
			return nil, err
		}
		dirStream = append(dirStream, sector...)
		sectorNum = int32(fat[sectorNum])
	}

	numDirs := len(dirStream) / dirEntrySize
	dirEntries := make([]dirEntry, numDirs)
	if err := binary.Read(bytes.NewReader(dirStream), binary.LittleEndian, &dirEntries); err != nil {
		return nil, err
	}

	return &Reader{r, fat, dirEntries}, nil
}

// ReadStream finds a stream by name and returns its content.
func (r *Reader) ReadStream(name string) ([]byte, error) {
	for _, entry := range r.dirEntries {
		if entry.ObjectType == 2 { // Stream Object
			entryName := utf16BytesToString(entry.Name, entry.NameLen)
			if entryName == name {
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
