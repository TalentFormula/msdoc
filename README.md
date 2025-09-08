# MsDoc

> An implementation of the .doc file format in go

## (Planned) Project structure
```
msdoc/
├── go.mod
├── README.md
├── cmd/
│   └── msdocdump/          # CLI tool for debugging/dumping DOC file info
├── pkg/
│   └── msdoc/              # Public API (users import this)
│       ├── doc.go          # Entry points (Open, Document struct)
│       ├── reader.go       # Text() and Metadata()
│       └── writer.go       # (stub for future Save support)
├── ole2/                   # OLE2 compound file parser
│   ├── reader.go
│   ├── sector.go
│   └── directory.go
├── fib/                    # File Information Block parsing
│   ├── fib.go
│   └── fib_types.go
├── streams/                # WordDocument, Table, Data streams
│   ├── worddocument.go
│   ├── table.go
│   └── data.go
├── structures/             # Shared DOC data structures (CP, PLC, FKP, etc.)
│   ├── cp.go
│   ├── plc.go
│   ├── fkp.go
│   └── sep.go
└── tests/
    ├── testdata/           # Sample .doc files
    ├── reader_test.go
    ├── fib_test.go
    └── ole2_test.go

```

## (Planned) Public High Level API

```go
package main

import (
	"fmt"
	"log"

	"github.com/TalentFormula/msdoc/pkg"
)

func main() {
	// Open a .doc file
	doc, err := msdoc.Open("sample.doc")
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
	fmt.Println("=== Metadata ===")
	fmt.Printf("Title: %s\n", meta.Title)
	fmt.Printf("Author: %s\n", meta.Author)
	fmt.Printf("Created: %s\n", meta.Created)
}
```

