package docx

import (
	"archive/zip"
	"io"
)

func copyOriginalFile(f *zip.File, zipWriter *zip.Writer) error {
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

func readFileContent(f *zip.File) ([]byte, error) {
	fileInZip, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer fileInZip.Close()

	return io.ReadAll(fileInZip)
}

func writeFile(filename string, zipWriter *zip.Writer, content []byte) error {
	newFile, err := zipWriter.CreateHeader(&zip.FileHeader{
		Name:   filename,
		Method: zip.Deflate,
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

func replaceFileContent(f *zip.File, zipWriter *zip.Writer, content []byte) error {
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
