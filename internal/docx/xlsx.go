package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"text/template"
)

func applyTemplateToCells(f *zip.File, templateValues any, fileContent []byte) ([]byte, error) {
	tmpl, err := template.New(f.Name).
		Funcs(template.FuncMap{
			"toNumberCell": toNumberCell,
		}).
		Parse(patchXML(string(fileContent)))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, templateValues); err != nil {
		return nil, fmt.Errorf("unable to execute template: %w", err)
	}

	return buf.Bytes(), nil
}

// ModifyXLSXInMemoryFromZipFile modifies an internal file inside an XLSX embedded in a zip.File.
// It returns a modified XLSX as []byte.
func ModifyXLSXInMemoryFromZipFile(xlsxFile *zip.File, templateValues any) ([]byte, error) {
	var sharedStringsNumbers map[int]string

	// Read XLSX zip into memory
	xlsxData, err := readFileContent(xlsxFile)
	if err != nil {
		return nil, fmt.Errorf("failed to read embedded XLSX file: %w", err)
	}

	// Create zip reader from in-memory XLSX data
	r, err := zip.NewReader(bytes.NewReader(xlsxData), int64(len(xlsxData)))
	if err != nil {
		return nil, fmt.Errorf("failed to open XLSX zip reader: %w", err)
	}

	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	found := false

	sharedStringsMatcher := regexp.MustCompile(`xl/(sharedStrings\d*)\.xml`)
	sheetNMatcher := regexp.MustCompile(`xl/worksheets/sheet\d*\.xml`)
	for _, f := range r.File {
		// avoid processing chartN files for the next for-loop
		if sheetNMatcher.MatchString(f.Name) {
			continue
		}

		fileContent, err := readFileContent(f)
		if err != nil {
			return nil, fmt.Errorf("error reading file %s: %w", f.Name, err)
		}

		w, err := zipWriter.CreateHeader(&f.FileHeader)
		if err != nil {
			return nil, fmt.Errorf("error creating file in zip: %w", err)
		}

		matchedSharedStrings := sharedStringsMatcher.MatchString(f.Name)
		if matchedSharedStrings {
			fileContent, err = applyTemplateToCells(f, templateValues, fileContent)
			if err != nil {
				return nil, fmt.Errorf("error applying template to file %s: %w", f.Name, err)
			}

			found = true
			fileContent, sharedStringsNumbers = GetReferencedSharedStringsByIndexAndCleanup(fileContent)
		}

		if _, err := w.Write(fileContent); err != nil {
			return nil, fmt.Errorf("error writing file %s: %w", f.Name, err)
		}
	}

	for _, f := range r.File {
		// avoid processing other files again
		if !sheetNMatcher.MatchString(f.Name) {
			continue
		}

		fileContent, err := readFileContent(f)
		if err != nil {
			return nil, fmt.Errorf("error reading file %s: %w", f.Name, err)
		}

		w, err := zipWriter.CreateHeader(&f.FileHeader)
		if err != nil {
			return nil, fmt.Errorf("error creating file in zip: %w", err)
		}

		matchedChartN := sheetNMatcher.MatchString(f.Name)
		if matchedChartN {
			var chartValues []string
			fileContent, chartValues, err = ReplaceSharedStringIndicesWithValues(fileContent, sharedStringsNumbers)
			if err != nil {
				return nil, fmt.Errorf("error replacing indexes in file %s: %w", f.Name, err)
			}

			XlsxFiles[xlsxFile.Name] = XlsxData{
				ChartNumbers: chartValues,
			}

			found = true
		}

		if _, err := w.Write(fileContent); err != nil {
			return nil, fmt.Errorf("error writing file %s: %w", f.Name, err)
		}
	}

	if !found {
		return nil, fmt.Errorf("internal file %s not found in embedded XLSX", sharedStringsMatcher.String())
	}

	if err := zipWriter.Close(); err != nil {
		return nil, fmt.Errorf("error closing zip writer: %w", err)
	}

	return buf.Bytes(), nil
}

func WriteXLSXIntoZip(f *zip.File, docxZipWriter *zip.Writer, templateValues any) error {
	//worksheets/sheet|
	xlsxBytes, err := ModifyXLSXInMemoryFromZipFile(f, templateValues)
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
