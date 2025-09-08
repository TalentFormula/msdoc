package structures

import "fmt"

// CP (Character Position) is an unsigned 32-bit integer that specifies
// a zero-based index of a character in the document text.
type CP uint32

// MaxCP is the maximum valid character position value.
const MaxCP = CP(0x7FFFFFFF)

// Common errors
var (
	ErrInvalidCP = fmt.Errorf("invalid character position")
)

// IsValid returns true if the CP value is within valid range.
func (cp CP) IsValid() bool {
	return cp <= MaxCP
}

// ToInt returns the CP as a regular int for array indexing.
// This is safe as long as the CP has been validated.
func (cp CP) ToInt() int {
	return int(cp)
}

// Distance calculates the number of characters between two CPs.
// Returns 0 if start >= end.
func (cp CP) Distance(end CP) uint32 {
	if cp >= end {
		return 0
	}
	return uint32(end - cp)
}
