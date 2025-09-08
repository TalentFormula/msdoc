package msdoc

// This file contains stubs for future write/save support.
// 
// The msdoc library currently supports reading .doc files only.
// Write support would require implementing:
//   - Document creation from scratch
//   - Text insertion and formatting
//   - Stream writing and OLE2 compound document creation
//   - FIB generation and piece table construction
//
// Example future API:
//
//   doc := msdoc.New()
//   doc.SetTitle("My Document")
//   doc.AddText("Hello, world!")
//   err := doc.Save("output.doc")
//
// This functionality is not yet implemented.
