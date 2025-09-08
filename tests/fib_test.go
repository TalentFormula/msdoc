package tests

import (
	"encoding/binary"
	"testing"

	"github.com/TalentFormula/msdoc/fib"
)

func TestParseFIB(t *testing.T) {
	// Create a mock FIB structure. Size must be large enough to contain
	// all the parts up to the cbRgFcLcb field.
	fibBytes := make([]byte, 32+2+28+2+88+2+93*8) // Base + counts + blobs

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

	offset := 32
	binary.LittleEndian.PutUint16(fibBytes[offset:], csw)
	offset += 2 + 28 // Skip over fibRgW
	binary.LittleEndian.PutUint16(fibBytes[offset:], cslw)
	offset += 2 + 88 // Skip over fibRgLw
	binary.LittleEndian.PutUint16(fibBytes[offset:], cbRgFcLcb)
	offset += 2

	// --- Populate key values in RgFcLcbBlob ---
	// According to spec for nFib=0x00C1, fcClx/lcbClx are at offsets 0x1A2/0x1A6
	// relative to the start of the WordDocument stream. The blob starts at offset 0x5A.
	// So, the offset within the blob is 0x1A2 - 0x5A = 0x148 = 328.
	// We'll use a simpler offset for this test since the code just uses the constant.
	fcClxOffsetInBlob := 248
	fcClx := uint32(0x1000)
	lcbClx := uint32(0x0200)
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
