package main

import (
	"fmt"
	"os"

	"github.com/TalentFormula/msdoc/ole2"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: liststreams <file>")
		os.Exit(1)
	}
	filePath := os.Args[1]
	file, err := os.Open(filePath)
	if err != nil {
		fmt.Printf("Failed to open file: %v\n", err)
		os.Exit(1)
	}
	defer file.Close()

	reader, err := ole2.NewReader(file)
	if err != nil {
		fmt.Printf("Failed to create OLE2 reader: %v\n", err)
		os.Exit(1)
	}

	streams := reader.ListStreams()
	fmt.Println("Streams found:")
	for _, s := range streams {
		fmt.Printf("- '%s'\n", s)
	}
}
