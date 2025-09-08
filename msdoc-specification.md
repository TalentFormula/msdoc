# Microsoft Word DOC File Format - Complete Technical Specification

## Document Information
- **Format Name**: Microsoft Word Binary File Format (.doc)
- **Specification Version**: MS-DOC 12.4 (August 2025)
- **Applicable Versions**: Microsoft Word 97, 2000, 2002 (XP), and 2003
- **File Extension**: .doc
- **MIME Type**: application/msword
- **File Signature**: D0 CF 11 E0 A1 B1 1A E1 (OLE2 Compound File)

---

## 1. Overview

The Microsoft Word DOC file format is a proprietary binary file format that stores word processing documents. The format is based on the OLE2 Compound File Binary Format (CFBF), which provides a file system-like structure within a single file.

### 1.1 Format Foundation
- **Base Technology**: OLE2 Compound Document Format
- **Structure**: Multi-stream hierarchical storage system
- **Data Organization**: Streams and storages within compound document
- **Maximum File Size**: ~2.1 GB (0x7FFFFFFF bytes per stream)

---

## 2. File Structure

### 2.1 OLE2 Compound Document Structure

#### 2.1.1 File Header
The DOC file begins with the standard OLE2 compound file header:
- **File Signature**: `D0 CF 11 E0 A1 B1 1A E1` (8 bytes)
- **Minor Version**: 0x003E (2 bytes)
- **Major Version**: 0x003E or 0x0004 (2 bytes)
- **Byte Order**: 0xFFFE (2 bytes)
- **Sector Size**: 0x0009 (512 bytes) or 0x000C (4096 bytes)
- **Mini Sector Size**: 0x0006 (64 bytes)

#### 2.1.2 Directory Structure
- **Root Entry**: Parent container for all streams and storages
- **Directory Stream**: Index of all streams and storages
- **Sector Allocation Tables**: Define sector positions within file

### 2.2 Core Streams

#### 2.2.1 WordDocument Stream (Required)
- **Purpose**: Main document content and File Information Block (FIB)
- **Location**: FIB starts at offset 0
- **Size Limit**: Maximum 0x7FFFFFFF bytes
- **Content**: Document text, formatting references, and structure information

#### 2.2.2 Table Stream (Required - Either 1Table or 0Table)
- **Selection**: Determined by `fWhichTblStm` flag in FIB
- **0Table Stream**: Used when `fWhichTblStm` = 0
- **1Table Stream**: Used when `fWhichTblStm` = 1
- **Purpose**: Contains formatting data and document structure information
- **Size Limit**: Maximum 0x7FFFFFFF bytes
- **Encryption**: For encrypted documents, begins with EncryptionHeader at offset 0

#### 2.2.3 Data Stream (Optional)
- **Purpose**: Additional document data referenced from FIB
- **Structure**: No predefined structure
- **Size Limit**: Maximum 0x7FFFFFFF bytes if present
- **Presence**: Only exists if referenced from other file components

---

## 3. File Information Block (FIB)

### 3.1 FIB Overview
The FIB is a variable-length master record that begins every WordDocument stream at offset 0. It contains:
- Document identification information
- File pointers to various document portions
- Formatting and structure references
- Version and compatibility information

### 3.2 FIB Structure Layout

```
FIB Structure:
┌─────────────────┐
│ FibBase (32)    │  Fixed-size base portion
├─────────────────┤
│ csw (2)         │  Count of 16-bit values in fibRgW
├─────────────────┤
│ fibRgW (28)     │  FibRgW97 structure
├─────────────────┤
│ cslw (2)        │  Count of 32-bit values in fibRgLw
├─────────────────┤
│ fibRgLw (88)    │  FibRgLw97 structure
├─────────────────┤
│ cbRgFcLcb (2)   │  Count of 64-bit values in fibRgFcLcbBlob
├─────────────────┤
│ fibRgFcLcbBlob  │  Variable-size FibRgFcLcb structure
├─────────────────┤
│ cswNew (2)      │  Count of 16-bit values in fibRgCswNew
├─────────────────┤
│ fibRgCswNew     │  Optional variable-size structure
└─────────────────┘
```

### 3.3 FibBase Structure (32 bytes)

#### 3.3.1 Field Definitions

| Offset | Size | Field | Description |
|--------|------|-------|-------------|
| 0 | 2 | wIdent | Word Binary File identifier (MUST be 0xA5EC) |
| 2 | 2 | nFib | File format version number (SHOULD be 0x00C1) |
| 4 | 2 | unused | Undefined value (MUST be ignored) |
| 6 | 2 | lid | Install language ID (LID) |
| 8 | 2 | pnNext | Offset to AutoText FIB (×512 bytes) |
| 10 | 2 | flags1 | Various boolean flags |
| 12 | 2 | nFibBack | Compatibility version (SHOULD be 0x00BF) |
| 14 | 4 | lKey | Encryption/obfuscation key |
| 18 | 1 | envr | Environment (MUST be 0) |
| 19 | 1 | flags2 | Additional boolean flags |
| 20 | 2 | reserved3 | Reserved (MUST be 0) |
| 22 | 2 | reserved4 | Reserved (MUST be 0) |
| 24 | 4 | reserved5 | Undefined (MUST be ignored) |
| 28 | 4 | reserved6 | Undefined (MUST be ignored) |

#### 3.3.2 Flags1 Bit Fields (Offset 10-11)

| Bit | Field | Description |
|-----|-------|-------------|
| 0 | fDot | Document template flag |
| 1 | fGlsy | AutoText-only document flag |
| 2 | fComplex | Incremental save operation flag |
| 3 | fHasPic | Pictures in document flag |
| 4-7 | cQuickSaves | Quick save count (4 bits) |
| 8 | fEncrypted | Document encrypted flag |
| 9 | fWhichTblStm | Table stream selection (0=0Table, 1=1Table) |
| 10 | fReadOnlyRecommended | Read-only mode recommended flag |
| 11 | fWriteReservation | Write-reservation password flag |
| 12 | fExtChar | Extended character flag (MUST be 1) |
| 13 | fLoadOverride | Language override flag |
| 14 | fFarEast | East Asian language flag |
| 15 | fObfuscated | XOR obfuscation flag |

#### 3.3.3 Flags2 Bit Fields (Offset 19)

| Bit | Field | Description |
|-----|-------|-------------|
| 0 | fMac | Mac compatibility (MUST be 0) |
| 1 | fEmptySpecial | Empty special (SHOULD be 0) |
| 2 | fLoadOverridePage | Page property override flag |
| 3 | reserved1 | Reserved bit |
| 4 | reserved2 | Reserved bit |
| 5-7 | fSpare0 | Spare bits (3 bits) |

### 3.4 Variable Sections

#### 3.4.1 Section Size Specifications

| nFib Value | cbRgFcLcb | cswNew |
|------------|-----------|--------|
| 0x00C1 | 0x005D | 0 |
| 0x00D9 | 0x006C | 0x0002 |
| 0x0101 | 0x0088 | 0x0002 |
| 0x010C | 0x00A4 | 0x0002 |
| 0x0112 | 0x00B7 | 0x0005 |

---

## 4. Data Structures

### 4.1 Character Position (CP)

#### 4.1.1 Definition
A Character Position (CP) is an unsigned 32-bit integer representing a zero-based index for characters in the document.

#### 4.1.2 Properties
- **Data Type**: Unsigned 32-bit integer
- **Range**: 0 to 0x7FFFFFFF (2,147,483,647)
- **Purpose**: Index characters, object anchors, and control characters
- **Note**: Consecutive CPs may not be at adjacent file locations

#### 4.1.3 Content Types
- Main document text
- Footnote text
- Header/footer text  
- Object anchors
- Control characters (paragraph marks, etc.)

### 4.2 PLC (Plex) Structure

#### 4.2.1 Overview
The PLC structure is an array format used throughout DOC files for organizing character positions with associated data.

#### 4.2.2 Structure Layout
```
PLC Structure:
┌─────────────────┐
│ CP[0]          │ ─┐
├─────────────────┤  │
│ CP[1]          │  │ Array of CPs
├─────────────────┤  │ (n+1 elements)
│ ...            │  │
├─────────────────┤  │
│ CP[n]          │ ─┘
├─────────────────┤
│ Data[0]        │ ─┐
├─────────────────┤  │
│ Data[1]        │  │ Array of data elements
├─────────────────┤  │ (n elements)
│ ...            │  │
├─────────────────┤  │
│ Data[n-1]      │ ─┘
└─────────────────┘
```

#### 4.2.3 Constraints
- Number of CPs MUST be one more than data elements
- CPs MUST appear in ascending order
- All data elements MUST be the same size
- Size calculation: n = (cbPlc - 4) / (cbData + 4)

---

## 5. Document Content Organization

### 5.1 Text Storage

#### 5.1.1 Location
Document text begins at position specified by `fib.fcMin` in the WordDocument stream.

#### 5.1.2 Content Types
- Main document body text
- Footnotes and endnotes
- Headers and footers
- Comments and annotations
- Control characters (paragraph marks, section breaks)
- Object anchors for embedded elements

### 5.2 Formatting Information

#### 5.2.1 FKP Structures
**Formatted Disk Pages (FKPs)** contain formatting information:
- **CHPs**: Character property formatting
- **PAPs**: Paragraph property formatting  
- **LVCs**: List property formatting

#### 5.2.2 Organization
- FKPs follow the document text
- Different FKP types are interleaved (not grouped)
- Each FKP contains multiple formatting records

### 5.3 Section Formatting
**SEPXs (Section Property eXtensions)** contain section-level formatting:
- Page setup information
- Headers and footers references
- Column layout specifications
- Page numbering settings

---

## 6. Encryption and Security

### 6.1 Encryption Methods

#### 6.1.1 XOR Obfuscation
- **Activation**: When fEncrypted=1 and fObfuscated=1
- **Key Location**: lKey field in FibBase
- **Method**: Simple XOR obfuscation algorithm
- **Purpose**: Basic content protection

#### 6.1.2 Standard Encryption
- **Activation**: When fEncrypted=1 and fObfuscated=0
- **Header Location**: Beginning of Table stream
- **Size**: Specified by lKey field in FibBase
- **Method**: Standard encryption algorithms

### 6.2 Password Protection
- **Write Reservation**: fWriteReservation flag
- **Read-Only Recommendation**: fReadOnlyRecommended flag
- **Document Protection**: Various document restriction settings

---

## 7. Version History and Compatibility

### 7.1 Format Versions

#### 7.1.1 Version Evolution
- **Word 1.0-2.0**: Early proprietary format (discontinued)
- **Word 6.0-95**: Modified binary format
- **Word 97-2003**: OLE2-based format (current specification)
- **Word 2007+**: DOCX default with DOC compatibility

#### 7.1.2 nFib Version Codes
- **0x00C1**: Word 97 format
- **0x00D9**: Word 2000 format  
- **0x0101**: Word 2002 format
- **0x010C**: Word 2003 format
- **0x0112**: Enhanced Word 2003 format

### 7.2 Compatibility Considerations

#### 7.2.1 Language Support
- East Asian language handling varies by version
- Spanish, German, French language recording rules
- Vietnamese, Thai, Hindi support in later versions

#### 7.2.2 Feature Support
- AutoText attachment capabilities
- Template functionality
- Embedded object support
- Macro support and security

---

## 8. Implementation Guidelines

### 8.1 Reading DOC Files

#### 8.1.1 Basic Steps
1. Validate OLE2 compound file structure
2. Locate and read WordDocument stream
3. Parse FIB structure starting at offset 0
4. Determine Table stream (0Table or 1Table)
5. Read document text from fcMin position
6. Parse formatting information using FIB pointers

#### 8.1.2 Error Handling
- Validate all structure sizes and boundaries
- Check for encryption before processing
- Handle version differences appropriately
- Verify stream existence before reading

### 8.2 Writing DOC Files

#### 8.2.1 Structure Creation
1. Create OLE2 compound document container
2. Write WordDocument stream with proper FIB
3. Create appropriate Table stream
4. Write document text and formatting data
5. Update all file pointers in FIB
6. Finalize compound document structure

#### 8.2.2 Compatibility Requirements
- Set appropriate nFib version number
- Include required structure elements
- Maintain proper stream size limits
- Handle language and encoding correctly

---

## 9. Technical Specifications Summary

### 9.1 File Limits
- **Maximum file size**: ~2.1 GB (theoretical)
- **Maximum stream size**: 0x7FFFFFFF bytes each
- **Maximum text length**: 0x7FFFFFFF characters
- **Sector size**: 512 or 4096 bytes

### 9.2 Data Types
- **CP**: Unsigned 32-bit integer (character position)
- **FC**: File position (32-bit unsigned integer)  
- **LCB**: Length in bytes (32-bit unsigned integer)
- **LID**: Language identifier (16-bit unsigned integer)

### 9.3 Required Streams
- **WordDocument**: Main document content (required)
- **0Table or 1Table**: Formatting data (required)
- **Data**: Additional document data (optional)

### 9.4 Optional Components
- **Summary Information**: Document properties
- **Document Summary Information**: Extended properties  
- **Object Storage**: Embedded OLE objects
- **VBA Storage**: Macro code and data

---

## 10. References and Resources

### 10.1 Official Documentation
- **Microsoft Open Specifications**: [MS-DOC] Word (.doc) Binary File Format
- **Version 12.4**: Published August 19, 2025
- **License**: Microsoft Open Specification Promise

### 10.2 Related Specifications  
- **[MS-CFB]**: Compound File Binary File Format
- **[MS-OLEDS]**: OLE1.0 and OLE2.0 Formats
- **[MS-OFFCRYPTO]**: Office Document Cryptography Structure

### 10.3 Implementation Tools
- **Microsoft Word 97-2003**: Reference implementation
- **Open Source Libraries**: wvWare, Apache POI, LibreOffice
- **Reverse Engineering**: Historical community efforts

---

*This specification document represents the complete technical details of the Microsoft Word DOC binary file format as documented in the official Microsoft specifications and related technical resources.*