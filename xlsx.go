package gotemplatedocx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"

	"github.com/JJJJJJack/go-template-docx/internal/xlsx"
	goziputils "github.com/JJJJJJack/go-zip-utils"
)

type chartData struct {
	chartNumbers []string
}

type xlsxChartsMap map[string]chartData

// modifyXlsxInMemoryFromZipFile modifies an internal file inside an XLSX embedded in a zip.File.
// It returns a modified XLSX as []byte.
func (dt *DocxTemplate) modifyXlsxInMemoryFromZipFile(xlsxFile *zip.File, templateValues any) ([]byte, error) {
	var sharedStringsNumbers map[int]string
	// key: old index, value: new index
	var sharedStringsNewIndexes map[int]int

	// Read XLSX zip into memory
	xlsxData, err := goziputils.ReadZipFileContent(xlsxFile)
	if err != nil {
		return nil, fmt.Errorf("failed to read embedded XLSX file: %w", err)
	}

	xlsxZipMap, err := goziputils.NewZipMapFromBytes(xlsxData)
	if err != nil {
		return nil, fmt.Errorf("failed to create XLSX zip map: %w", err)
	}

	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	// Copy all files except the ones that will be processed
	sheetNMatcher := regexp.MustCompile(`xl/worksheets/sheet\d*\.xml`)
	sharedStringsMatcher := regexp.MustCompile(`xl/(sharedStrings\d*)\.xml`)
	sharedStringsFilename := "xl/sharedStrings.xml"

	for filename, f := range xlsxZipMap {
		switch {
		case
			sheetNMatcher.MatchString(filename),
			sharedStringsMatcher.MatchString(filename):
			continue
		}

		err := goziputils.CopyFile(zipWriter, f)
		if err != nil {
			return nil, fmt.Errorf("unable to copy original embedding xlsx file '%s': %w", f.Name, err)
		}
	}

	// work on sharedStrings.xml
	sharedStringsFile := xlsxZipMap[sharedStringsFilename]
	if sharedStringsFile == nil {
		return nil, fmt.Errorf("shared strings file '%s' not found in embedded XLSX", sharedStringsFilename)
	}

	sharedStringsContent, err := goziputils.ReadZipFileContent(sharedStringsFile)
	if err != nil {
		return nil, fmt.Errorf("error reading file '%s': %w", sharedStringsFile.Name, err)
	}

	sharedStringsContent, err = xlsx.ApplyTemplateToCells(sharedStringsFile, templateValues, sharedStringsContent)
	if err != nil {
		return nil, fmt.Errorf("error applying template to file '%s': %w", sharedStringsFile.Name, err)
	}

	sharedStringsContent, sharedStringsNumbers, sharedStringsNewIndexes, err = xlsx.GetReferencedSharedStringsByIndexAndCleanup(sharedStringsContent)
	if err != nil {
		return nil, fmt.Errorf("error cleaning up shared strings in file '%s': %w", sharedStringsFile.Name, err)
	}

	sharedStringsCount := uint(0)
	for i := 1; ; i++ {
		sheetN := fmt.Sprintf("xl/worksheets/sheet%d.xml", i)

		f := xlsxZipMap[sheetN]
		if f == nil {
			break
		}

		fileContent, err := goziputils.ReadZipFileContent(f)
		if err != nil {
			return nil, fmt.Errorf("error reading zip file content '%s': %w", f.Name, err)
		}

		var chartValues []string
		fileContent, chartValues, err = xlsx.UpdateSheet(fileContent, sharedStringsNumbers, sharedStringsNewIndexes)
		if err != nil {
			return nil, fmt.Errorf("error replacing shared strings indexes in file '%s': %w", f.Name, err)
		}

		dt.xlsxChartsMeta[xlsxFile.Name] = chartData{
			chartNumbers: chartValues,
		}

		sharedStringsRefs, err := xlsx.GetCountFromXml(fileContent)
		if err != nil {
			return nil, fmt.Errorf("error getting shared strings refs count from file '%s': %w", f.Name, err)
		}

		sharedStringsCount += sharedStringsRefs

		err = goziputils.RewriteFileIntoZipWriter(zipWriter, f, fileContent)
		if err != nil {
			return nil, fmt.Errorf("error writing file '%s': %w", f.Name, err)
		}
	}

	// need to be here, after all sheets have been processed we know the real count
	sharedStringsContent, err = xlsx.RecountSharedStringsCountAndUniqueCountAttributes(sharedStringsContent, sharedStringsCount)
	if err != nil {
		return nil, fmt.Errorf("error recounting sharedStrings file '%s': %w", sharedStringsFile.Name, err)
	}

	err = goziputils.RewriteFileIntoZipWriter(zipWriter, sharedStringsFile, sharedStringsContent)
	if err != nil {
		return nil, fmt.Errorf("error writing sharedStrings file '%s': %w", sharedStringsFile.Name, err)
	}

	if err := zipWriter.Close(); err != nil {
		return nil, fmt.Errorf("error closing zip writer: %w", err)
	}

	return buf.Bytes(), nil
}

func (dt *DocxTemplate) writeXlsxIntoZip(f *zip.File, docxZipWriter *zip.Writer, templateValues any) error {
	xlsxBytes, err := dt.modifyXlsxInMemoryFromZipFile(f, templateValues)
	if err != nil {
		return fmt.Errorf("error modifying XLSX in memory: %w", err)
	}

	err = goziputils.RewriteFileIntoZipWriter(docxZipWriter, f, xlsxBytes)
	if err != nil {
		return fmt.Errorf("error creating entry in zip: %w", err)
	}

	return nil
}
