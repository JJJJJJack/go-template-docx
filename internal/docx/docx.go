package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path"
	"regexp"
	"slices"
	"strings"
)

type Template struct {
	templateFile string
	media        []media
	output       bytes.Buffer
	isApplied    bool
	rel          *relationship
	relMedia     []mediaRel
}

func NewTemplate(templateFile string) *Template {
	return &Template{
		templateFile: templateFile,
		output:       bytes.Buffer{},
		isApplied:    false,
		media:        []media{},
		rel:          &relationship{},
		relMedia:     []mediaRel{},
	}
}

func (t *Template) Media(filename string, data []byte) {
	t.media = append(t.media, media{
		filename: filename,
		data:     data,
	})
}

func (t *Template) Apply(templateValues any) error {
	zipWriter := zip.NewWriter(&t.output)

	r, err := zip.OpenReader(t.templateFile)
	if err != nil {
		return fmt.Errorf("unable to open template file %s: %w", t.templateFile, err)
	}
	defer r.Close()

	for _, m := range t.media {
		filename := path.Join("word/media", m.filename)
		err = writeFile(filename, zipWriter, m.data)
		if err != nil {
			return fmt.Errorf("unable to write media file %s: %w", filename, err)
		}
	}

	ctFile := "[Content_Types].xml"
	relFile := "word/_rels/document.xml.rels"
	for _, f := range r.File {
		if f.Name != relFile {
			continue
		}

		relData, err := readFileContent(f)
		if err != nil {
			fmt.Println("unable to read rel file")
			break
		}

		t.rel, err = parseRelationship(relData)
		if err != nil {
			fmt.Println("unable to parse relationship file:", err)
		}

		break
	}

	toSkip := []string{
		relFile,
		ctFile,
	}

	chartRelToTargetXlsx := make(map[string]string)
	chartsRelMatcher := regexp.MustCompile(`word/charts/_rels/chart\d*?\.xml.rels`)
	chartsMatcher := regexp.MustCompile(`word/charts/chart\d*?\.xml`)
	xlsxMatcher := regexp.MustCompile(`/embeddings/Microsoft_Excel_Worksheet\d*?\.xlsx`)
	headerFooterDocumentMatcher := regexp.MustCompile(`word/(header|footer|document)\d*?\.xml`)
	for _, f := range r.File {
		if slices.Contains(toSkip, f.Name) {
			continue
		}

		matchedChartsRel := chartsRelMatcher.MatchString(f.Name)
		if matchedChartsRel {
			fileContent, err := readFileContent(f)
			if err != nil {
				return fmt.Errorf("unable to read chart rel file %s: %w", f.Name, err)
			}

			chartsRelationships, _ := parseRelationship(fileContent)
			for _, relationship := range chartsRelationships.Relationships {
				if !xlsxMatcher.MatchString(relationship.Target) {
					continue
				}

				formattedTarget := strings.Replace(relationship.Target, "../", "word/", 1)
				chartNumber, err := ExtractChartName(f.Name)
				if err != nil {
					return fmt.Errorf("unable to extract chart name from file %s: %w", f.Name, err)
				}
				chartRelToTargetXlsx[chartNumber] = formattedTarget
			}
		}

		matchedChart := chartsMatcher.MatchString(f.Name)
		if matchedChart {
			continue
		}

		matchedXlsx := xlsxMatcher.MatchString(f.Name)
		if matchedXlsx {
			err := WriteXLSXIntoZip(f, zipWriter, templateValues)
			if err != nil {
				return fmt.Errorf("unable to write XLSX file %s: %w", f.Name, err)
			}
			continue
		}

		// I don't know how many headers/footers there are, so I use a regex
		matchedHeaderFooterDocument := headerFooterDocumentMatcher.MatchString(f.Name)
		if !matchedHeaderFooterDocument {
			err = copyOriginalFile(f, zipWriter)
			if err != nil {
				return fmt.Errorf("unable to copy original file %s: %w", f.Name, err)
			}
			continue
		}

		media, err := applyTemplate(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to file %s: %w", f.Name, err)
		}

		t.relMedia = append(t.relMedia, media...)
	}

	for _, f := range r.File {
		matchedChart := chartsMatcher.MatchString(f.Name)
		if !matchedChart {
			continue
		}

		fileContent, err := readFileContent(f)
		if err != nil {
			return fmt.Errorf("unable to read chart file %s: %w", f.Name, err)
		}

		fileContent, err = ApplyTemplateToChart(f, templateValues, fileContent)
		if err != nil {
			return fmt.Errorf("unable to apply template to file %s: %w", f.Name, err)
		}

		chartFileNumber, err := ExtractChartName(f.Name)
		if err != nil {
			return fmt.Errorf("unable to extract chart name from file %s: %w", f.Name, err)
		}

		xlsxFileTarget := chartRelToTargetXlsx[chartFileNumber]
		fileContent, err = UpdateDocxChartWithXlsxChartValues(fileContent, XlsxFiles[xlsxFileTarget].ChartNumbers)
		if err != nil {
			return fmt.Errorf("unable to replace preview zeros in chart file %s: %w", f.Name, err)
		}

		w, err := zipWriter.CreateHeader(&f.FileHeader)
		if err != nil {
			return fmt.Errorf("error creating file in zip: %w", err)
		}

		_, err = w.Write(fileContent)
		if err != nil {
			return fmt.Errorf("error writing file %s: %w", f.Name, err)
		}
	}

	if len(t.relMedia) != 0 {
		t.rel.addMediaToRels(t.relMedia)
		for _, f := range r.File {
			if f.Name != relFile {
				continue
			}

			xmlContent, err := t.rel.toXML()
			if err != nil {
				fmt.Println("unable to marshal rels:", err)
				break
			}

			err = replaceFileContent(f, zipWriter, []byte(xmlContent))
			if err != nil {
				fmt.Println("unable to replace rel file:", err)
				break
			}
		}
		for _, f := range r.File {
			if f.Name != ctFile {
				continue
			}

			ctData, err := readFileContent(f)
			if err != nil {
				fmt.Println("unable to read content types file")
				break
			}

			parsedCt, err := parseContentTypes(ctData)
			if err != nil {
				fmt.Println("unable to parse content types file:", err)
			}

			parsedCt.ensureImageDefaults("png", "image/png")
			updatedCt, err := parsedCt.toXML()
			if err != nil {
				fmt.Println("unable to marshal rels:", err)
				break
			}

			err = replaceFileContent(f, zipWriter, []byte(updatedCt))
			if err != nil {
				fmt.Println("unable to replace rel file:", err)
				break
			}
			break
		}
	}

	err = zipWriter.Close()
	if err != nil {
		return fmt.Errorf("unable to close zip writer: %w", err)
	}

	t.isApplied = true
	return nil
}

func (t *Template) Save(output string) error {
	if !t.isApplied {
		return fmt.Errorf("template not applied")
	}

	t.isApplied = false
	return os.WriteFile(output, t.output.Bytes(), 0644)
}

func (t *Template) Bytes() ([]byte, error) {
	if !t.isApplied {
		return nil, fmt.Errorf("template not applied")
	}

	t.isApplied = false
	return t.output.Bytes(), nil
}
