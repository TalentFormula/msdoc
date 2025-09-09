package tests

import (
	"testing"

	"github.com/TalentFormula/msdoc/pkg"
)

// TestMetadataExtraction tests metadata extraction from sample documents
func TestMetadataExtraction(t *testing.T) {
	testCases := []struct {
		filename           string
		expectedTitle      string
		expectedAuthor     string
		expectedSubject    string
		expectedKeywords   string
		expectedComments   string
		expectedAppName    string
		expectedCompany    string
		expectedManager    string
		expectedContentStatus string
		expectedContentType string
		expectedCategory   string
	}{
		{
			filename: "testdata/sample-2.doc", 
			expectedTitle: "The title is working", 
			expectedAuthor: "Advik B",
			expectedSubject: "",
			expectedKeywords: "",
			expectedComments: "NO",
			expectedAppName: "Microsoft Office Word",
		},
		{
			filename: "testdata/sample-3.doc",
			expectedTitle: "The Third Title",
			expectedAuthor: "",
			expectedSubject: "TalentSort", 
			expectedKeywords: "tag1",
			expectedComments: "Yayy",
			expectedAppName: "Microsoft Office Word",
			expectedCompany: "TalentFormula",
			expectedManager: "Who Knows",
			expectedContentStatus: "ready",
			expectedContentType: "application/msword",
			expectedCategory: "dumb",
		},
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
			t.Logf("  Company: '%s'", metadata.Company)
			t.Logf("  Manager: '%s'", metadata.Manager)
			t.Logf("  Category: '%s'", metadata.Category)
			t.Logf("  ContentStatus: '%s'", metadata.ContentStatus)
			t.Logf("  ContentType: '%s'", metadata.ContentType)
			
			// Check if title matches expected
			if metadata.Title != tc.expectedTitle {
				t.Errorf("Expected title '%s', got '%s'", tc.expectedTitle, metadata.Title)
			}
			
			// Check if author matches expected
			if metadata.Author != tc.expectedAuthor {
				t.Errorf("Expected author '%s', got '%s'", tc.expectedAuthor, metadata.Author)
			}
			
			// Check subject if specified
			if tc.expectedSubject != "" && metadata.Subject != tc.expectedSubject {
				t.Errorf("Expected subject '%s', got '%s'", tc.expectedSubject, metadata.Subject)
			}
			
			// Check keywords if specified
			if tc.expectedKeywords != "" && metadata.Keywords != tc.expectedKeywords {
				t.Errorf("Expected keywords '%s', got '%s'", tc.expectedKeywords, metadata.Keywords)
			}
			
			// Check comments if specified
			if tc.expectedComments != "" && metadata.Comments != tc.expectedComments {
				t.Errorf("Expected comments '%s', got '%s'", tc.expectedComments, metadata.Comments)
			}
			
			// Check app name if specified
			if tc.expectedAppName != "" && metadata.ApplicationName != tc.expectedAppName {
				t.Errorf("Expected application name '%s', got '%s'", tc.expectedAppName, metadata.ApplicationName)
			}
			
			// Check company if specified
			if tc.expectedCompany != "" && metadata.Company != tc.expectedCompany {
				t.Errorf("Expected company '%s', got '%s'", tc.expectedCompany, metadata.Company)
			}
			
			// Check manager if specified
			if tc.expectedManager != "" && metadata.Manager != tc.expectedManager {
				t.Errorf("Expected manager '%s', got '%s'", tc.expectedManager, metadata.Manager)
			}
			
			// Check content status if specified
			if tc.expectedContentStatus != "" && metadata.ContentStatus != tc.expectedContentStatus {
				t.Errorf("Expected content status '%s', got '%s'", tc.expectedContentStatus, metadata.ContentStatus)
			}
			
			// Check content type if specified
			if tc.expectedContentType != "" && metadata.ContentType != tc.expectedContentType {
				t.Errorf("Expected content type '%s', got '%s'", tc.expectedContentType, metadata.ContentType)
			}
			
			// Check category if specified
			if tc.expectedCategory != "" && metadata.Category != tc.expectedCategory {
				t.Errorf("Expected category '%s', got '%s'", tc.expectedCategory, metadata.Category)
			}
		})
	}
}