package structures

// PLC (Plex) is a common structure in .doc files. It is an array of
// Character Positions (CPs) followed by an array of data elements.
// The number of CPs is always one more than the number of data elements.
type PLC struct {
	CPs  []CP
	Data [][]byte // Generic representation of data elements
}
