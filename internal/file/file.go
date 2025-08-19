package file

import (
	"errors"
	"fmt"
	"os"
)

// FindFirstMissingFile returns the first file path that does not exist.
// If all files exist, it returns an error.
func FindFirstMissingFile(filenames []string) (string, error) {
	for _, filename := range filenames {
		if _, err := os.Stat(filename); os.IsNotExist(err) {
			return filename, nil
		} else if err != nil {
			return "", fmt.Errorf("error checking %q: %w", filename, err)
		}
	}
	return "", errors.New("all files exist")
}
