// Package objects provides support for extracting embedded objects from .doc files.
//
// This package handles OLE objects, images, charts, and other embedded content
// stored within Word documents according to the MS-DOC specification.
package objects

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"

	"github.com/TalentFormula/msdoc/ole2"
)

// ObjectType represents the type of embedded object.
type ObjectType int

const (
	ObjectTypeUnknown  ObjectType = iota
	ObjectTypeOLE                 // OLE object (Excel sheet, PowerPoint, etc.)
	ObjectTypeImage               // Image (BMP, PNG, JPEG, etc.)
	ObjectTypeChart               // Chart or graph
	ObjectTypeEquation            // Mathematical equation
	ObjectTypeDrawing             // Drawing or shape
)

// EmbeddedObject represents an object embedded in the document.
type EmbeddedObject struct {
	Type      ObjectType // Type of the embedded object
	Name      string     // Object name or description
	ClassName string     // OLE class name (for OLE objects)
	Data      []byte     // Raw object data
	IconData  []byte     // Icon representation data
	Size      int64      // Size of the object data
	Position  uint32     // Position in document where object is referenced
	IsLinked  bool       // True if object is linked rather than embedded
	LinkPath  string     // Path to linked file (if applicable)
}

// ObjectPool manages embedded objects within a .doc file.
type ObjectPool struct {
	reader  *ole2.Reader
	objects map[uint32]*EmbeddedObject
}

// NewObjectPool creates a new ObjectPool for the given OLE2 reader.
func NewObjectPool(reader *ole2.Reader) *ObjectPool {
	return &ObjectPool{
		reader:  reader,
		objects: make(map[uint32]*EmbeddedObject),
	}
}

// LoadObjects loads all embedded objects from the ObjectPool stream.
func (op *ObjectPool) LoadObjects() error {
	// Try to read the ObjectPool stream
	poolData, err := op.reader.ReadStream("ObjectPool")
	if err != nil {
		// ObjectPool stream may not exist if there are no embedded objects
		return nil
	}

	return op.parseObjectPool(poolData)
}

// parseObjectPool parses the ObjectPool stream data.
func (op *ObjectPool) parseObjectPool(data []byte) error {
	reader := bytes.NewReader(data)

	for reader.Len() > 0 {
		obj, err := op.parseObject(reader)
		if err != nil {
			if err == io.EOF {
				break
			}
			return fmt.Errorf("failed to parse object: %w", err)
		}

		if obj != nil {
			op.objects[obj.Position] = obj
		}
	}

	return nil
}

// parseObject parses a single embedded object from the stream.
func (op *ObjectPool) parseObject(reader *bytes.Reader) (*EmbeddedObject, error) {
	// Read object header
	var header struct {
		Signature uint32 // Object signature
		Size      uint32 // Object size
		Type      uint16 // Object type
		Flags     uint16 // Object flags
	}

	if err := binary.Read(reader, binary.LittleEndian, &header); err != nil {
		return nil, err
	}

	// Validate object signature
	if header.Signature != 0x00000501 { // Standard OLE object signature
		return nil, fmt.Errorf("invalid object signature: 0x%x", header.Signature)
	}

	obj := &EmbeddedObject{
		Size:     int64(header.Size),
		Position: uint32(reader.Size()) - uint32(reader.Len()), // Current position
	}

	// Determine object type
	obj.Type = op.determineObjectType(header.Type, header.Flags)
	obj.IsLinked = (header.Flags & 0x0001) != 0

	// Read object data
	if header.Size > 0 {
		obj.Data = make([]byte, header.Size)
		if _, err := reader.Read(obj.Data); err != nil {
			return nil, fmt.Errorf("failed to read object data: %w", err)
		}

		// Parse object-specific data
		if err := op.parseObjectData(obj); err != nil {
			return nil, fmt.Errorf("failed to parse object-specific data: %w", err)
		}
	}

	return obj, nil
}

// determineObjectType determines the object type from type and flags.
func (op *ObjectPool) determineObjectType(objType uint16, flags uint16) ObjectType {
	switch objType {
	case 0x0002: // OLE object
		return ObjectTypeOLE
	case 0x0003: // Static image
		return ObjectTypeImage
	case 0x0005: // Chart
		return ObjectTypeChart
	case 0x0007: // Equation
		return ObjectTypeEquation
	case 0x0008: // Drawing
		return ObjectTypeDrawing
	default:
		return ObjectTypeUnknown
	}
}

// parseObjectData parses type-specific object data.
func (op *ObjectPool) parseObjectData(obj *EmbeddedObject) error {
	if len(obj.Data) < 4 {
		return nil // Not enough data to parse
	}

	reader := bytes.NewReader(obj.Data)

	switch obj.Type {
	case ObjectTypeOLE:
		return op.parseOLEObject(obj, reader)
	case ObjectTypeImage:
		return op.parseImageObject(obj, reader)
	case ObjectTypeChart:
		return op.parseChartObject(obj, reader)
	default:
		// For unknown types, just store the raw data
		return nil
	}
}

// parseOLEObject parses OLE object data.
func (op *ObjectPool) parseOLEObject(obj *EmbeddedObject, reader *bytes.Reader) error {
	// Read OLE object header
	var oleHeader struct {
		Version uint32 // OLE version
		Flags   uint32 // OLE flags
		NameLen uint32 // Length of class name
	}

	if err := binary.Read(reader, binary.LittleEndian, &oleHeader); err != nil {
		return fmt.Errorf("failed to read OLE header: %w", err)
	}

	// Read class name
	if oleHeader.NameLen > 0 {
		nameBytes := make([]byte, oleHeader.NameLen)
		if _, err := reader.Read(nameBytes); err != nil {
			return fmt.Errorf("failed to read OLE class name: %w", err)
		}
		obj.ClassName = string(nameBytes)
	}

	// The remaining data is the actual OLE object
	remaining := make([]byte, reader.Len())
	reader.Read(remaining)
	obj.Data = remaining

	return nil
}

// parseImageObject parses image object data.
func (op *ObjectPool) parseImageObject(obj *EmbeddedObject, reader *bytes.Reader) error {
	// Read image header
	var imgHeader struct {
		Format uint32 // Image format (BMP, PNG, etc.)
		Width  uint32 // Image width
		Height uint32 // Image height
	}

	if err := binary.Read(reader, binary.LittleEndian, &imgHeader); err != nil {
		return fmt.Errorf("failed to read image header: %w", err)
	}

	// Determine image format
	obj.Name = op.getImageFormatName(imgHeader.Format)

	// The remaining data is the actual image
	remaining := make([]byte, reader.Len())
	reader.Read(remaining)
	obj.Data = remaining

	return nil
}

// parseChartObject parses chart object data.
func (op *ObjectPool) parseChartObject(obj *EmbeddedObject, reader *bytes.Reader) error {
	// Chart objects are typically Excel chart objects
	obj.ClassName = "Excel.Chart"
	obj.Name = "Chart"

	// The data is the chart definition
	remaining := make([]byte, reader.Len())
	reader.Read(remaining)
	obj.Data = remaining

	return nil
}

// getImageFormatName returns the image format name from format code.
func (op *ObjectPool) getImageFormatName(format uint32) string {
	switch format {
	case 0x216: // CF_BITMAP
		return "BMP"
	case 0x8000: // Custom PNG
		return "PNG"
	case 0x8001: // Custom JPEG
		return "JPEG"
	case 0x8002: // Custom GIF
		return "GIF"
	default:
		return "Unknown"
	}
}

// GetObject returns the embedded object at the given position.
func (op *ObjectPool) GetObject(position uint32) *EmbeddedObject {
	return op.objects[position]
}

// GetAllObjects returns all embedded objects.
func (op *ObjectPool) GetAllObjects() map[uint32]*EmbeddedObject {
	return op.objects
}

// ExtractObject extracts an embedded object and returns its data.
func (op *ObjectPool) ExtractObject(position uint32) (*EmbeddedObject, error) {
	obj := op.objects[position]
	if obj == nil {
		return nil, fmt.Errorf("no object found at position %d", position)
	}

	return obj, nil
}

// SaveObject saves an embedded object to a file.
func (obj *EmbeddedObject) SaveObject(filename string) error {
	if len(obj.Data) == 0 {
		return errors.New("no object data to save")
	}

	// Implementation would write obj.Data to filename
	// This is a placeholder for the actual file writing logic
	return fmt.Errorf("save functionality not yet implemented")
}

// GetObjectInfo returns human-readable information about the object.
func (obj *EmbeddedObject) GetObjectInfo() string {
	typeStr := obj.getTypeString()

	info := fmt.Sprintf("Type: %s", typeStr)
	if obj.Name != "" {
		info += fmt.Sprintf(", Name: %s", obj.Name)
	}
	if obj.ClassName != "" {
		info += fmt.Sprintf(", Class: %s", obj.ClassName)
	}
	info += fmt.Sprintf(", Size: %d bytes", obj.Size)
	if obj.IsLinked {
		info += fmt.Sprintf(", Linked to: %s", obj.LinkPath)
	}

	return info
}

// getTypeString returns a string representation of the object type.
func (obj *EmbeddedObject) getTypeString() string {
	switch obj.Type {
	case ObjectTypeOLE:
		return "OLE Object"
	case ObjectTypeImage:
		return "Image"
	case ObjectTypeChart:
		return "Chart"
	case ObjectTypeEquation:
		return "Equation"
	case ObjectTypeDrawing:
		return "Drawing"
	default:
		return "Unknown"
	}
}
