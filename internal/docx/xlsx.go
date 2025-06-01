package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"text/template"
)

func applyTemplateToCells(f *zip.File, data any, fileContent []byte) ([]byte, error) {
	tmpl, err := template.New(f.Name).
		Parse(string(patchXML(string(fileContent))))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, data); err != nil {
		return nil, fmt.Errorf("unable to execute template: %w", err)
	}

	return buf.Bytes(), nil
}

// ModifyXLSXInMemoryFromZipFile modifies an internal file inside an XLSX embedded in a zip.File.
// It returns a modified XLSX as []byte.
func ModifyXLSXInMemoryFromZipFile(xlsxFile *zip.File, fileMatcher string, data any) ([]byte, error) {
	// Open embedded XLSX file (it's itself a zip archive)
	xlsxReader, err := xlsxFile.Open()
	if err != nil {
		return nil, fmt.Errorf("failed to open embedded XLSX file: %w", err)
	}
	defer xlsxReader.Close()

	// Read XLSX zip into memory
	xlsxData, err := io.ReadAll(xlsxReader)
	if err != nil {
		return nil, fmt.Errorf("failed to read XLSX zip: %w", err)
	}

	// Create zip reader from in-memory XLSX data
	r, err := zip.NewReader(bytes.NewReader(xlsxData), int64(len(xlsxData)))
	if err != nil {
		return nil, fmt.Errorf("failed to open XLSX zip reader: %w", err)
	}

	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	found := false

	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return nil, fmt.Errorf("error opening file %s: %w", f.Name, err)
		}
		xmlContent, err := io.ReadAll(rc)

		rc.Close()
		if err != nil {
			return nil, fmt.Errorf("error reading file %s: %w", f.Name, err)
		}

		w, err := zipWriter.CreateHeader(&f.FileHeader)
		if err != nil {
			return nil, fmt.Errorf("error creating file in zip: %w", err)
		}

		matched, err := regexp.MatchString(fileMatcher, f.Name)
		if err != nil {
			return nil, fmt.Errorf("error matching file name %s with pattern %s: %w", f.Name, fileMatcher, err)
		}

		if matched {
			xmlContent, err = applyTemplateToCells(f, data, xmlContent)
			// mydebug.FindAndPrintSnippet(string(xmlContent), "categoryTest")
			// os.WriteFile("test.xml", xmlContent, 0644)
			found = true
		}

		if _, err := w.Write(xmlContent); err != nil {
			return nil, fmt.Errorf("error writing file %s: %w", f.Name, err)
		}
	}

	if !found {
		return nil, fmt.Errorf("internal file %s not found in embedded XLSX", fileMatcher)
	}

	if err := zipWriter.Close(); err != nil {
		return nil, fmt.Errorf("error closing zip writer: %w", err)
	}

	return buf.Bytes(), nil
}

func WriteXLSXIntoZip(docxZipWriter *zip.Writer, f *zip.File, data any) error {
	xlsxBytes, err := ModifyXLSXInMemoryFromZipFile(f, `xl/sharedStrings\d*.xml`, data)
	if err != nil {
		return fmt.Errorf("error modifying XLSX in memory: %w", err)
	}

	w, err := docxZipWriter.Create(f.Name)
	if err != nil {
		return fmt.Errorf("error creating entry in zip: %w", err)
	}

	_, err = w.Write(xlsxBytes)
	return err
}
