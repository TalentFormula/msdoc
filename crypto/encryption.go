package crypto

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
)

// EncryptionHeader represents the encryption information stored in the table stream
// for encrypted Word documents.
type EncryptionHeader struct {
	Version         uint16   // Encryption version
	EncryptionFlags uint32   // Encryption flags
	HeaderSize      uint32   // Size of encryption header
	ProviderType    uint32   // Cryptographic provider type
	AlgID           uint32   // Algorithm identifier
	AlgHashID       uint32   // Hash algorithm identifier
	KeySize         uint32   // Key size in bits
	ProviderName    string   // Cryptographic provider name
	Salt            []byte   // Random salt for key derivation
	EncryptedVerifier []byte // Encrypted verifier for password validation
	VerifierHash    []byte   // Hash of the verifier
}

// ParseEncryptionHeader parses the encryption header from table stream data.
func ParseEncryptionHeader(data []byte) (*EncryptionHeader, error) {
	if len(data) < 32 {
		return nil, errors.New("encryption header too small")
	}

	reader := bytes.NewReader(data)
	header := &EncryptionHeader{}

	// Read basic header fields
	if err := binary.Read(reader, binary.LittleEndian, &header.Version); err != nil {
		return nil, fmt.Errorf("failed to read version: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.EncryptionFlags); err != nil {
		return nil, fmt.Errorf("failed to read flags: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.HeaderSize); err != nil {
		return nil, fmt.Errorf("failed to read header size: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.ProviderType); err != nil {
		return nil, fmt.Errorf("failed to read provider type: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.AlgID); err != nil {
		return nil, fmt.Errorf("failed to read algorithm ID: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.AlgHashID); err != nil {
		return nil, fmt.Errorf("failed to read hash algorithm ID: %w", err)
	}

	if err := binary.Read(reader, binary.LittleEndian, &header.KeySize); err != nil {
		return nil, fmt.Errorf("failed to read key size: %w", err)
	}

	// Skip reserved fields and read provider name
	reader.Seek(8, 1) // Skip 8 bytes of reserved fields

	// Read provider name (null-terminated Unicode string)
	providerNameBytes := make([]byte, 64) // Maximum provider name length
	reader.Read(providerNameBytes)
	header.ProviderName = parseUnicodeString(providerNameBytes)

	// Read salt (16 bytes)
	header.Salt = make([]byte, 16)
	if _, err := reader.Read(header.Salt); err != nil {
		return nil, fmt.Errorf("failed to read salt: %w", err)
	}

	// Read encrypted verifier (16 bytes)
	header.EncryptedVerifier = make([]byte, 16)
	if _, err := reader.Read(header.EncryptedVerifier); err != nil {
		return nil, fmt.Errorf("failed to read encrypted verifier: %w", err)
	}

	// Read verifier hash (16 bytes for MD5)
	header.VerifierHash = make([]byte, 16)
	if _, err := reader.Read(header.VerifierHash); err != nil {
		return nil, fmt.Errorf("failed to read verifier hash: %w", err)
	}

	return header, nil
}

// IsRC4Encryption returns true if the encryption uses RC4 algorithm.
func (h *EncryptionHeader) IsRC4Encryption() bool {
	// RC4 algorithm ID
	return h.AlgID == 0x6801 // CALG_RC4
}

// IsPasswordProtected returns true if the document is password protected.
func (h *EncryptionHeader) IsPasswordProtected() bool {
	return len(h.EncryptedVerifier) > 0 && len(h.VerifierHash) > 0
}

// ValidatePassword checks if the provided password is correct for this document.
func (h *EncryptionHeader) ValidatePassword(password string) (bool, error) {
	if !h.IsPasswordProtected() {
		return false, errors.New("document is not password protected")
	}

	// Generate decryption key from password and salt
	key, err := GenerateDecryptionKey(password, h.Salt)
	if err != nil {
		return false, fmt.Errorf("failed to generate key: %w", err)
	}

	// Create RC4 cipher
	rc4, err := NewRC4(key)
	if err != nil {
		return false, fmt.Errorf("failed to create RC4 cipher: %w", err)
	}

	// Decrypt the verifier
	decryptedVerifier := rc4.Decrypt(h.EncryptedVerifier)

	// Hash the decrypted verifier
	verifierHash := GeneratePasswordHash(string(decryptedVerifier))

	// Compare with stored hash
	for i := 0; i < len(verifierHash) && i < len(h.VerifierHash); i++ {
		if verifierHash[i] != h.VerifierHash[i] {
			return false, nil // Password is incorrect
		}
	}

	return true, nil
}

// CreateDecryptionCipher creates an RC4 cipher for decrypting document content.
func (h *EncryptionHeader) CreateDecryptionCipher(password string) (*RC4, error) {
	if !h.IsPasswordProtected() {
		return nil, errors.New("document is not password protected")
	}

	// Validate password first
	valid, err := h.ValidatePassword(password)
	if err != nil {
		return nil, err
	}
	if !valid {
		return nil, errors.New("incorrect password")
	}

	// Generate decryption key
	key, err := GenerateDecryptionKey(password, h.Salt)
	if err != nil {
		return nil, fmt.Errorf("failed to generate key: %w", err)
	}

	// Create RC4 cipher
	return NewRC4(key)
}

// parseUnicodeString extracts a null-terminated Unicode string from byte data.
func parseUnicodeString(data []byte) string {
	var result []rune
	for i := 0; i < len(data)-1; i += 2 {
		if data[i] == 0 && data[i+1] == 0 {
			break // Null terminator found
		}
		char := uint16(data[i]) | (uint16(data[i+1]) << 8)
		if char != 0 {
			result = append(result, rune(char))
		}
	}
	return string(result)
}