package fib

import (
    "bytes"
    "encoding/binary"
    "errors"
    "fmt"
)

// ParseFIB reads a byte slice (from the WordDocument stream)
// and parses it into a FileInformationBlock struct.
func ParseFIB(data []byte) (*FileInformationBlock, error) {
    if len(data) < 32 { // Minimum size for FibBase
        return nil, errors.New("fib: data too short for FibBase")
    }

    r := bytes.NewReader(data)
    fib := &FileInformationBlock{}

    // Read the fixed-size FibBase
    if err := binary.Read(r, binary.LittleEndian, &fib.Base); err != nil {
        return nil, fmt.Errorf("fib: could not read FibBase: %w", err)
    }

    // Validate Word document identifier
    if fib.Base.WIdent != 0xA5EC {
        return nil, errors.New("fib: invalid wIdent, not a Word document")
    }

    // Read counts for variable sections
    var csw, cslw, cbRgFcLcb, _ uint16
    binary.Read(r, binary.LittleEndian, &csw)
    r.Seek(28, 1) // Skip fibRgW for now
    binary.Read(r, binary.LittleEndian, &cslw)
    r.Seek(88, 1) // Skip fibRgLw for now
    binary.Read(r, binary.LittleEndian, &cbRgFcLcb)

    // Reset reader and parse structs properly
    r.Seek(0, 0)
    br := bytes.NewReader(data)

    if err := binary.Read(br, binary.LittleEndian, &fib.Base); err != nil {
        return nil, err
    }
    if err := binary.Read(br, binary.LittleEndian, &fib.Csw); err != nil {
        return nil, err
    }
    if err := binary.Read(br, binary.LittleEndian, &fib.FibRgW); err != nil {
        return nil, err
    }
    if err := binary.Read(br, binary.LittleEndian, &fib.Cslw); err != nil {
        return nil, err
    }
    if err := binary.Read(br, binary.LittleEndian, &fib.FibRgLw); err != nil {
        return nil, err
    }
    if err := binary.Read(br, binary.LittleEndian, &fib.CbRgFcLcb); err != nil {
        return nil, err
    }

    // Read the variable-length FibRgFcLcb
    fib.RgFcLcbBlob = make([]byte, fib.CbRgFcLcb*8)
    if err := binary.Read(br, binary.LittleEndian, &fib.RgFcLcbBlob); err != nil {
        return nil, fmt.Errorf("fib: could not read RgFcLcbBlob: %w", err)
    }

    // Parse known fields from RgFcLcbBlob for text extraction
    // A full parser would map this whole blob to a struct.
    // For now, we extract only what's needed.
    // fcClx is at offset 0x01A2 - 0x5A = 248 bytes into the blob for nFib=0x00C1
    const fcClxOffset = 248
    if len(fib.RgFcLcbBlob) >= fcClxOffset+8 {
        fib.RgFcLcb.FcClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset:])
        fib.RgFcLcb.LcbClx = binary.LittleEndian.Uint32(fib.RgFcLcbBlob[fcClxOffset+4:])
    }

    return fib, nil
}
