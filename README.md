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