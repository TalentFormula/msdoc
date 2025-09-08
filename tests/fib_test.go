package tests

import (
	"encoding/binary"
	"github.com/TalentFormula/msdoc/fib"
	"testing"
)

func TestParseFIB(t *testing.T) {
	// Create a mock FIB structure. Size must be large enough to contain
	// all the parts up to the cbRgFcLcb field.
	blobSizeInBytes := 93 * 8
	fibBytes := make([]byte, 32+2+28+2+80+2+blobSizeInBytes) // Base + counts + blobs (using binary.Read sizes)

	// --- Populate FibBase (first 32 bytes) ---
	wIdent := uint16(0xA5EC)
	nFib := uint16(0x00C1)
	flags1 := uint16(0x0200) // fWhichTblStm = 1 (use 1Table)

	binary.LittleEndian.PutUint16(fibBytes[0:], wIdent)
	binary.LittleEndian.PutUint16(fibBytes[2:], nFib)
	binary.LittleEndian.PutUint16(fibBytes[10:], flags1)

	// --- Populate Counts ---
	csw := uint16(14)       // Size of fibRgW in uint16
	cslw := uint16(22)      // Size of fibRgLw in uint32
	cbRgFcLcb := uint16(93) // Corresponds to nFib 0x00C1 (0x5D)

	offset := 32 // FibBase size when packed
	binary.LittleEndian.PutUint16(fibBytes[offset:], csw)
	offset += 2 + 28 // Skip over fibRgW  
	binary.LittleEndian.PutUint16(fibBytes[offset:], cslw)
	offset += 2 + 80 // Skip over fibRgLw (actual binary.Read size is 80, not 88)
	binary.LittleEndian.PutUint16(fibBytes[offset:], cbRgFcLcb)  
	offset += 2 // Offset is now at the start of the blob

	// --- Populate key values in RgFcLcbBlob ---
	// For Word 97 (nFib=0x00C1), fcClx/lcbClx are at offset 264 within the blob.
	// THIS IS THE CORRECTED VALUE.
	fcClxOffsetInBlob := 264
	fcClx := uint32(4096)
	lcbClx := uint32(512)
	binary.LittleEndian.PutUint32(fibBytes[offset+fcClxOffsetInBlob:], fcClx)
	binary.LittleEndian.PutUint32(fibBytes[offset+fcClxOffsetInBlob+4:], lcbClx)

	// --- Run the test ---
	parsedFIB, err := fib.ParseFIB(fibBytes)
	if err != nil {
		t.Fatalf("ParseFIB failed: %v", err)
	}

	if parsedFIB.Base.WIdent != wIdent {
		t.Errorf("Expected wIdent 0x%X, got 0x%X", wIdent, parsedFIB.Base.WIdent)
	}

	if parsedFIB.Base.NFib != nFib {
		t.Errorf("Expected nFib 0x%X, got 0x%X", nFib, parsedFIB.Base.NFib)
	}

	if parsedFIB.Base.Flags1 != flags1 {
		t.Errorf("Expected flags1 0x%X, got 0x%X", flags1, parsedFIB.Base.Flags1)
	}

	if parsedFIB.RgFcLcb.FcClx != fcClx {
		t.Errorf("Expected parsed FcClx %d, got %d", fcClx, parsedFIB.RgFcLcb.FcClx)
	}

	if parsedFIB.RgFcLcb.LcbClx != lcbClx {
		t.Errorf("Expected parsed LcbClx %d, got %d", lcbClx, parsedFIB.RgFcLcb.LcbClx)
	}
}
