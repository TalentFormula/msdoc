package tests

import (
	"testing"

	"github.com/TalentFormula/msdoc/pkg"
)

// TestMetadataExtraction tests metadata extraction from sample documents
func TestMetadataExtraction(t *testing.T) {
	testCases := []struct {
		filename      string
		expectedTitle string
		expectedAuthor string
	}{
		{"testdata/sample-2.doc", "The title is working", "Advik B"},
	}

	for _, tc := range testCases {
		t.Run(tc.filename, func(t *testing.T) {
			doc, err := msdoc.Open(tc.filename)
			if err != nil {
				t.Fatalf("Failed to open document: %v", err)
			}
			defer doc.Close()

			metadata := doc.Metadata()
			
			t.Logf("Extracted metadata:")
			t.Logf("  Title: '%s'", metadata.Title)
			t.Logf("  Author: '%s'", metadata.Author)
			t.Logf("  Subject: '%s'", metadata.Subject)
			t.Logf("  Keywords: '%s'", metadata.Keywords)
			t.Logf("  Comments: '%s'", metadata.Comments)
			t.Logf("  Template: '%s'", metadata.Template)
			t.Logf("  LastAuthor: '%s'", metadata.LastAuthor)
			t.Logf("  ApplicationName: '%s'", metadata.ApplicationName)
			
			// Check if title matches expected
			if metadata.Title != tc.expectedTitle {
				t.Errorf("Expected title '%s', got '%s'", tc.expectedTitle, metadata.Title)
			}
			
			// Check if author matches expected
			if metadata.Author != tc.expectedAuthor {
				t.Errorf("Expected author '%s', got '%s'", tc.expectedAuthor, metadata.Author)
			}
		})
	}
}