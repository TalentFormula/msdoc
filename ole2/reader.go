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
	err := binary.Read(io.NewSectionReader(r, 0, 76), binary.LittleEndian, &header)
	if err != nil {
		return nil, err
	}
	if header.Signature != headerSignature {
		return nil, errors.New("ole2: invalid signature")
	}

	// Read DIFAT
	difat := make([]int32, 109)
	err = binary.Read(io.NewSectionReader(r, 76, 436), binary.LittleEndian, &difat)
	if err != nil {
		return nil, err
	}

	// Read FAT sectors
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
	err = binary.Read(bytes.NewReader(fatSectors), binary.LittleEndian, &fat)
	if err != nil {
		return nil, err
	}

	// Read directory stream
	var dirStream []byte
	sectorNum := header.DirStartSector
	for sectorNum >= 0 {
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
	err = binary.Read(bytes.NewReader(dirStream), binary.LittleEndian, &dirEntries)
	if err != nil {
		return nil, err
	}

	return &Reader{r, fat, dirEntries}, nil
}

// ReadStream finds a stream by name and returns its content.
func (r *Reader) ReadStream(name string) ([]byte, error) {
	for _, entry := range r.dirEntries {
		if entry.ObjectType == 2 { // Stream Object
			entryName := utf16BytesToString(entry.Name[:], entry.NameLen)
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
				// Trim to actual size
				return streamData[:entry.StreamSize], nil
			}
		}
	}
	return nil, fmt.Errorf("ole2: stream '%s' not found", name)
}

// utf16BytesToString converts a UTF-16 byte sequence from a directory entry to a Go string.
func utf16BytesToString(b []uint16, nameLen uint16) string {
	if nameLen == 0 {
		return ""
	}
	// Name length includes the null terminator, so we subtract it.
	end := (nameLen / 2) - 1
	if end > uint16(len(b)) {
		end = uint16(len(b))
	}
	return string(utf16.Decode(b[:end]))
}
