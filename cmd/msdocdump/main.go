package main

import (
	"fmt"
	"log"
	"os"

	"github.com/TalentFormula/msdoc/pkg"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: msdocdump <file.doc>")
		os.Exit(1)
	}
	filename := os.Args[1]

	// Open the .doc file using our library
	doc, err := msdoc.Open(filename)
	if err != nil {
		log.Fatalf("failed to open DOC: %v", err)
	}
	defer doc.Close()

	// Extract text with markdown formatting (hyperlinks as [text](url))
	text, err := doc.MarkdownText()
	if err != nil {
		log.Fatalf("failed to extract text: %v", err)
	}
	fmt.Println("=== Document Text ===")
	fmt.Println(text)

	// Extract metadata
	meta := doc.Metadata()
	fmt.Println("\n=== Metadata ===")
	fmt.Printf("Title: %s\n", meta.Title)
	fmt.Printf("Subject: %s\n", meta.Subject)
	fmt.Printf("Author: %s\n", meta.Author)
	fmt.Printf("Keywords: %s\n", meta.Keywords)
	fmt.Printf("Comments: %s\n", meta.Comments)
	fmt.Printf("Application Name: %s\n", meta.ApplicationName)
	fmt.Printf("Company: %s\n", meta.Company)
	fmt.Printf("Manager: %s\n", meta.Manager)
	fmt.Printf("Category: %s\n", meta.Category)
	fmt.Printf("Content Status: %s\n", meta.ContentStatus)
	fmt.Printf("Content Type: %s\n", meta.ContentType)
	fmt.Printf("Created: %s\n", meta.Created)
}
