package utils

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
)

type ZipMap map[string]*zip.File

func NewZipMap(data []byte) (ZipMap, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("failed to create zip reader: %w", err)
	}

	zipMap := make(ZipMap)
	for _, f := range r.File {
		zipMap[f.Name] = f
	}

	return zipMap, nil
}

func CopyOriginalFile(f *zip.File, zipWriter *zip.Writer) error {
	fileInZip, err := f.Open()
	if err != nil {
		return err
	}
	defer fileInZip.Close()

	newFile, err := zipWriter.CreateHeader(&zip.FileHeader{
		Name:   f.Name,
		Method: f.FileHeader.Method,
	})
	if err != nil {
		return err
	}
	_, err = io.Copy(newFile, fileInZip)
	if err != nil {
		return err
	}

	return nil
}

// ReadZipFileContent reads the content of a file in a zip archive.
func ReadZipFileContent(f *zip.File) ([]byte, error) {
	fileInZip, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer fileInZip.Close()

	return io.ReadAll(fileInZip)
}

// RewriteFileIntoZipWriter recreates a file with the zipWriter with new
// content and sets the FileHeader accordingly.
func RewriteFileIntoZipWriter(f *zip.File, zipWriter *zip.Writer, content []byte) error {
	newHeader := f.FileHeader

	newHeader.UncompressedSize64 = uint64(len(content))
	newHeader.CompressedSize64 = 0

	newFile, err := zipWriter.CreateHeader(&newHeader)
	if err != nil {
		return err
	}

	_, err = newFile.Write(content)
	if err != nil {
		return err
	}

	return nil
}

// ZipWriteFile creates a new file in the zip archive with the given filename
func ZipWriteFile(filename string, zipWriter *zip.Writer, content []byte) error {
	newFile, err := zipWriter.CreateHeader(&zip.FileHeader{
		Name:               filename,
		Method:             zip.Store,
		UncompressedSize64: uint64(len(content)),
	})
	if err != nil {
		return err
	}

	_, err = newFile.Write(content)
	if err != nil {
		return err
	}

	return nil
}

func ReplaceFileContent(f *zip.File, zipWriter *zip.Writer, content []byte) error {
	fileInZip, err := f.Open()
	if err != nil {
		return err
	}
	defer fileInZip.Close()

	newFile, err := zipWriter.CreateHeader(&zip.FileHeader{
		Name:   f.Name,
		Method: f.FileHeader.Method,
	})
	if err != nil {
		return err
	}

	_, err = newFile.Write(content)
	if err != nil {
		return err
	}

	return nil
}

// ExtractChartFilename now only works with a single submatch
func ExtractChartFilename(path string) (string, error) {
	re := regexp.MustCompile(`(chart\d+)\.xml`)
	matches := re.FindStringSubmatch(path)
	if len(matches) < 2 {
		return "", fmt.Errorf("no chart name found")
	}
	return matches[1], nil
}
