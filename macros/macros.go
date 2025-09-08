// Package macros provides support for extracting and processing VBA macros from .doc files.
//
// This package handles the Macros stream and VBA project structures according to
// the MS-DOC specification and related OLE2/VBA documentation.
package macros

import (
	"bytes"
	"compress/zlib"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"strings"

	"github.com/TalentFormula/msdoc/ole2"
)

// VBAProject represents a VBA project contained in the document.
type VBAProject struct {
	Name        string            // Project name
	Description string            // Project description
	HelpFile    string            // Help file path
	Modules     map[string]*Module // VBA modules by name
	References  []*Reference      // External references
	Protected   bool              // True if project is protected
	Password    string            // Project password (if known)
}

// Module represents a VBA module (code module, class module, or form).
type Module struct {
	Name         string     // Module name
	Type         ModuleType // Module type
	Code         string     // VBA source code
	Compressed   bool       // True if code is compressed
	StreamName   string     // Storage stream name
	Offset       uint32     // Offset within stream
	Size         uint32     // Uncompressed size
}

// Reference represents an external reference used by the VBA project.
type Reference struct {
	Name        string // Reference name
	Description string // Reference description
	GUID        string // Reference GUID
	Version     string // Reference version
	Path        string // Reference file path
}

// ModuleType represents the type of VBA module.
type ModuleType int

const (
	ModuleStandard ModuleType = iota // Standard code module
	ModuleClass                      // Class module
	ModuleForm                       // UserForm module
	ModuleDocument                   // Document module (ThisDocument)
)

// MacroExtractor handles extraction of VBA macros from .doc files.
type MacroExtractor struct {
	reader *ole2.Reader
}

// NewMacroExtractor creates a new macro extractor for the given OLE2 reader.
func NewMacroExtractor(reader *ole2.Reader) *MacroExtractor {
	return &MacroExtractor{
		reader: reader,
	}
}

// HasMacros checks if the document contains VBA macros.
func (me *MacroExtractor) HasMacros() bool {
	// Check for Macros storage
	_, err := me.reader.ReadStream("Macros")
	if err == nil {
		return true
	}

	// Check for VBA storage (alternative location)
	_, err = me.reader.ReadStream("_VBA_PROJECT")
	return err == nil
}

// ExtractProject extracts the complete VBA project from the document.
func (me *MacroExtractor) ExtractProject() (*VBAProject, error) {
	if !me.HasMacros() {
		return nil, errors.New("document does not contain VBA macros")
	}

	project := &VBAProject{
		Modules:    make(map[string]*Module),
		References: make([]*Reference, 0),
	}

	// Try to read project information
	if err := me.parseProjectInfo(project); err != nil {
		return nil, fmt.Errorf("failed to parse project info: %w", err)
	}

	// Extract modules
	if err := me.extractModules(project); err != nil {
		return nil, fmt.Errorf("failed to extract modules: %w", err)
	}

	return project, nil
}

// parseProjectInfo parses the project-level information.
func (me *MacroExtractor) parseProjectInfo(project *VBAProject) error {
	// Read dir stream for project information
	dirData, err := me.reader.ReadStream("Macros/dir")
	if err != nil {
		// Try alternative location
		dirData, err = me.reader.ReadStream("_VBA_PROJECT")
		if err != nil {
			return fmt.Errorf("failed to read project directory: %w", err)
		}
	}

	return me.parseDirStream(project, dirData)
}

// parseDirStream parses the dir stream containing project metadata.
func (me *MacroExtractor) parseDirStream(project *VBAProject, data []byte) error {
	reader := bytes.NewReader(data)

	for reader.Len() > 0 {
		// Read record header
		var recordType uint16
		if err := binary.Read(reader, binary.LittleEndian, &recordType); err != nil {
			if err == io.EOF {
				break
			}
			return fmt.Errorf("failed to read record type: %w", err)
		}

		var recordLength uint32
		if err := binary.Read(reader, binary.LittleEndian, &recordLength); err != nil {
			return fmt.Errorf("failed to read record length: %w", err)
		}

		// Read record data
		recordData := make([]byte, recordLength)
		if _, err := reader.Read(recordData); err != nil {
			return fmt.Errorf("failed to read record data: %w", err)
		}

		// Process record based on type
		switch recordType {
		case 0x01: // Project information
			me.parseProjectRecord(project, recordData)
		case 0x07: // Module information
			module, err := me.parseModuleRecord(recordData)
			if err != nil {
				return fmt.Errorf("failed to parse module record: %w", err)
			}
			if module != nil {
				project.Modules[module.Name] = module
			}
		case 0x0D: // Reference information
			ref, err := me.parseReferenceRecord(recordData)
			if err != nil {
				return fmt.Errorf("failed to parse reference record: %w", err)
			}
			if ref != nil {
				project.References = append(project.References, ref)
			}
		}
	}

	return nil
}

// parseProjectRecord parses project-level information.
func (me *MacroExtractor) parseProjectRecord(project *VBAProject, data []byte) {
	// Extract project name, description, etc.
	reader := bytes.NewReader(data)
	
	// Read null-terminated strings
	project.Name = me.readNullTerminatedString(reader)
	project.Description = me.readNullTerminatedString(reader)
	project.HelpFile = me.readNullTerminatedString(reader)
}

// parseModuleRecord parses module information.
func (me *MacroExtractor) parseModuleRecord(data []byte) (*Module, error) {
	if len(data) < 8 {
		return nil, errors.New("module record too short")
	}

	reader := bytes.NewReader(data)
	module := &Module{}

	// Read module name
	module.Name = me.readNullTerminatedString(reader)
	
	// Read module type
	var moduleType uint32
	if err := binary.Read(reader, binary.LittleEndian, &moduleType); err != nil {
		return nil, fmt.Errorf("failed to read module type: %w", err)
	}
	module.Type = ModuleType(moduleType)

	// Read stream name
	module.StreamName = me.readNullTerminatedString(reader)

	// Read offset and size
	if err := binary.Read(reader, binary.LittleEndian, &module.Offset); err != nil {
		return nil, fmt.Errorf("failed to read module offset: %w", err)
	}
	
	if err := binary.Read(reader, binary.LittleEndian, &module.Size); err != nil {
		return nil, fmt.Errorf("failed to read module size: %w", err)
	}

	return module, nil
}

// parseReferenceRecord parses reference information.
func (me *MacroExtractor) parseReferenceRecord(data []byte) (*Reference, error) {
	reader := bytes.NewReader(data)
	ref := &Reference{}

	// Read reference information
	ref.Name = me.readNullTerminatedString(reader)
	ref.Description = me.readNullTerminatedString(reader)
	ref.GUID = me.readNullTerminatedString(reader)
	ref.Version = me.readNullTerminatedString(reader)
	ref.Path = me.readNullTerminatedString(reader)

	return ref, nil
}

// extractModules extracts the actual VBA code for all modules.
func (me *MacroExtractor) extractModules(project *VBAProject) error {
	for _, module := range project.Modules {
		if err := me.extractModuleCode(module); err != nil {
			return fmt.Errorf("failed to extract code for module %s: %w", module.Name, err)
		}
	}
	return nil
}

// extractModuleCode extracts the VBA source code for a specific module.
func (me *MacroExtractor) extractModuleCode(module *Module) error {
	// Read the module stream
	streamPath := fmt.Sprintf("Macros/%s", module.StreamName)
	streamData, err := me.reader.ReadStream(streamPath)
	if err != nil {
		// Try alternative VBA location
		streamPath = fmt.Sprintf("VBA/%s", module.StreamName)
		streamData, err = me.reader.ReadStream(streamPath)
		if err != nil {
			return fmt.Errorf("failed to read module stream %s: %w", module.StreamName, err)
		}
	}

	// Extract code from the specified offset
	if uint32(len(streamData)) < module.Offset {
		return fmt.Errorf("stream too short for module offset %d", module.Offset)
	}

	codeData := streamData[module.Offset:]
	if module.Size > 0 && uint32(len(codeData)) > module.Size {
		codeData = codeData[:module.Size]
	}

	// Check if code is compressed
	if len(codeData) > 0 && codeData[0] == 0x01 {
		// Compressed VBA code
		module.Compressed = true
		decompressed, err := me.decompressVBACode(codeData[1:]) // Skip compression flag
		if err != nil {
			return fmt.Errorf("failed to decompress VBA code: %w", err)
		}
		module.Code = string(decompressed)
	} else {
		// Uncompressed code
		module.Compressed = false
		module.Code = string(codeData)
	}

	return nil
}

// decompressVBACode decompresses VBA source code using the custom VBA compression algorithm.
func (me *MacroExtractor) decompressVBACode(data []byte) ([]byte, error) {
	if len(data) == 0 {
		return nil, errors.New("no data to decompress")
	}

	// Check for zlib compression first
	if len(data) > 2 && data[0] == 0x78 {
		reader, err := zlib.NewReader(bytes.NewReader(data))
		if err == nil {
			defer reader.Close()
			decompressed, err := io.ReadAll(reader)
			if err == nil {
				return decompressed, nil
			}
		}
	}

	// Custom VBA decompression algorithm
	return me.decompressVBACustom(data)
}

// decompressVBACustom implements the custom VBA compression algorithm.
func (me *MacroExtractor) decompressVBACustom(data []byte) ([]byte, error) {
	// This is a simplified implementation of VBA decompression
	// A complete implementation would handle the full VBA compression algorithm
	
	var output bytes.Buffer
	reader := bytes.NewReader(data)

	for reader.Len() > 0 {
		b, err := reader.ReadByte()
		if err != nil {
			break
		}

		if b >= 0x20 && b <= 0x7E {
			// Printable ASCII character
			output.WriteByte(b)
		} else if b == 0x0D {
			// Carriage return
			output.WriteByte('\r')
		} else if b == 0x0A {
			// Line feed
			output.WriteByte('\n')
		} else if b == 0x09 {
			// Tab
			output.WriteByte('\t')
		}
		// Skip other control characters
	}

	return output.Bytes(), nil
}

// readNullTerminatedString reads a null-terminated string from the reader.
func (me *MacroExtractor) readNullTerminatedString(reader *bytes.Reader) string {
	var result strings.Builder
	
	for {
		b, err := reader.ReadByte()
		if err != nil || b == 0 {
			break
		}
		result.WriteByte(b)
	}
	
	return result.String()
}

// GetModuleCode returns the VBA code for a specific module.
func (project *VBAProject) GetModuleCode(moduleName string) (string, bool) {
	module, exists := project.Modules[moduleName]
	if !exists {
		return "", false
	}
	return module.Code, true
}

// GetAllModuleNames returns the names of all modules in the project.
func (project *VBAProject) GetAllModuleNames() []string {
	names := make([]string, 0, len(project.Modules))
	for name := range project.Modules {
		names = append(names, name)
	}
	return names
}

// HasMacroFunctions checks if any module contains macro functions.
func (project *VBAProject) HasMacroFunctions() bool {
	for _, module := range project.Modules {
		if strings.Contains(module.Code, "Sub ") || 
		   strings.Contains(module.Code, "Function ") {
			return true
		}
	}
	return false
}

// GetModuleByType returns all modules of the specified type.
func (project *VBAProject) GetModulesByType(moduleType ModuleType) []*Module {
	var modules []*Module
	for _, module := range project.Modules {
		if module.Type == moduleType {
			modules = append(modules, module)
		}
	}
	return modules
}

// String returns a string representation of the module type.
func (mt ModuleType) String() string {
	switch mt {
	case ModuleStandard:
		return "Standard"
	case ModuleClass:
		return "Class"
	case ModuleForm:
		return "Form"
	case ModuleDocument:
		return "Document"
	default:
		return "Unknown"
	}
}