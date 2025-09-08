package main

import (
    "fmt"
    "github.com/TalentFormula/msdoc/pkg/msdoc"
    "log"
    "os"
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

    // Extract plain text
    text, err := doc.Text()
    if err != nil {
        log.Fatalf("failed to extract text: %v", err)
    }
    fmt.Println("=== Document Text ===")
    fmt.Println(text)

    // Extract metadata
    meta := doc.Metadata()
    fmt.Println("\n=== Metadata ===")
    fmt.Printf("Title: %s\n", meta.Title)
    fmt.Printf("Author: %s\n", meta.Author)
    fmt.Printf("Created: %s\n", meta.Created)
}
