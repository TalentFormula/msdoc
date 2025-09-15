// Package crypto provides cryptographic functions for .doc file decryption.
//
// This package implements the RC4 encryption/decryption algorithm used by
// Microsoft Word documents for password protection and encryption.
package crypto

import (
	"crypto/md5"
	"errors"
	"fmt"
)

// RC4 represents an RC4 cipher context.
type RC4 struct {
	s    [256]byte
	i, j byte
}

// NewRC4 creates a new RC4 cipher with the given key.
func NewRC4(key []byte) (*RC4, error) {
	if len(key) == 0 {
		return nil, errors.New("rc4: key cannot be empty")
	}

	rc4 := &RC4{}

	// Key-scheduling algorithm (KSA)
	for i := 0; i < 256; i++ {
		rc4.s[i] = byte(i)
	}

	j := byte(0)
	for i := 0; i < 256; i++ {
		j = j + rc4.s[i] + key[i%len(key)]
		rc4.s[i], rc4.s[j] = rc4.s[j], rc4.s[i]
	}

	rc4.i = 0
	rc4.j = 0

	return rc4, nil
}

// Decrypt decrypts the given data using RC4.
// RC4 is symmetric, so this function can also be used for encryption.
func (rc4 *RC4) Decrypt(data []byte) []byte {
	output := make([]byte, len(data))

	for k := 0; k < len(data); k++ {
		rc4.i++
		rc4.j += rc4.s[rc4.i]
		rc4.s[rc4.i], rc4.s[rc4.j] = rc4.s[rc4.j], rc4.s[rc4.i]
		output[k] = data[k] ^ rc4.s[rc4.s[rc4.i]+rc4.s[rc4.j]]
	}

	return output
}

// GeneratePasswordHash creates a password hash compatible with Word documents.
// This implements the Word 97-2003 password hashing algorithm.
func GeneratePasswordHash(password string) []byte {
	if len(password) == 0 {
		return nil
	}

	// Convert password to UTF-16LE
	utf16Password := make([]byte, 0, len(password)*2)
	for _, r := range password {
		utf16Password = append(utf16Password, byte(r), byte(r>>8))
	}

	// Generate MD5 hash
	hash := md5.Sum(utf16Password)
	return hash[:]
}

// GenerateDecryptionKey creates the decryption key from password and document salt.
// This follows the MS-DOC specification for password-based encryption.
func GenerateDecryptionKey(password string, salt []byte) ([]byte, error) {
	if len(password) == 0 {
		return nil, errors.New("password cannot be empty")
	}

	if len(salt) < 16 {
		return nil, fmt.Errorf("salt must be at least 16 bytes, got %d", len(salt))
	}

	// Generate password hash
	passwordHash := GeneratePasswordHash(password)

	// Combine password hash with document salt
	combined := append(passwordHash, salt[:16]...)

	// Generate final key hash
	finalHash := md5.Sum(combined)
	return finalHash[:], nil
}

// VerifyPassword checks if the given password matches the document's password hash.
func VerifyPassword(password string, expectedHash []byte, salt []byte) (bool, error) {
	if len(expectedHash) != 16 {
		return false, errors.New("expected hash must be 16 bytes")
	}

	key, err := GenerateDecryptionKey(password, salt)
	if err != nil {
		return false, err
	}

	// Compare the generated key with expected hash
	for i := 0; i < 16; i++ {
		if key[i] != expectedHash[i] {
			return false, nil
		}
	}

	return true, nil
}
