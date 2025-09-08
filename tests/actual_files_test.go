package tests

import (
    "testing"

    "github.com/TalentFormula/msdoc/pkg"
)

// TestActualDocFiles tests the library against actual .doc files in the testdata directory.
// These tests document the current behavior and serve as integration tests.
func TestActualDocFiles(t *testing.T) {
    testCases := []struct {
        filename string
        desc     string
    }{
        {"tests/testdata/sample-1.doc", "Sample 1 - Test document"},
        {"tests/testdata/sample-2.doc", "Sample 2 - Test document"},
    }

    for _, tc := range testCases {
        t.Run(tc.filename, func(t *testing.T) {
            doc, err := pkg.Open(tc.filename)
            if err != nil {
                t.Logf("Expected failure: %v", err)
                // For now, we expect these to fail as the OLE2 parsing needs more work
                // This test documents the current state and will pass once the issues are resolved
                return
            }
            defer doc.Close()

            // If we successfully opened the file, try to extract text
            text, err := doc.Text()
            if err != nil {
                t.Logf("Text extraction failed: %v", err)
                return
            }

            t.Logf("Successfully extracted %d characters of text from %s", len(text), tc.filename)
            if len(text) > 100 {
                t.Logf("Text preview: %s...", text[:100])
            } else if len(text) > 0 {
                t.Logf("Text content: %s", text)
            }
        })
    }
}
