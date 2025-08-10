package gotemplatedocx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path"
	"regexp"
	"slices"
	"strings"
	"time"

	"github.com/JJJJJJack/go-template-docx/internal/docx"
	"github.com/JJJJJJack/go-template-docx/internal/file"
	"github.com/JJJJJJack/go-template-docx/internal/utils"
	"github.com/JJJJJJack/go-template-docx/internal/xlsx"
)

type docxTemplate struct {
	outputFilename string
	bytes          []byte
	reader         *zip.Reader
	output         bytes.Buffer
	rel            *docx.Relationship
	relMedia       []docx.MediaRel
	media          []docx.Media
}

func NewDocxTemplateFromBytes(docxBytes []byte) (*docxTemplate, error) {
	bytesReader := bytes.NewReader(docxBytes)
	if bytesReader == nil {
		return nil, fmt.Errorf("unable to create bytes reader for DOCX file")
	}

	docxReader, err := zip.NewReader(bytesReader, int64(len(docxBytes)))
	if err != nil {
		return nil, fmt.Errorf("unable to create reader for DOCX file: %v", err)
	}

	return &docxTemplate{
		outputFilename: "",
		bytes:          docxBytes,
		reader:         docxReader,
		output:         bytes.Buffer{},
		media:          []docx.Media{},
		rel:            &docx.Relationship{},
		relMedia:       []docx.MediaRel{},
	}, nil
}

func NewDocxTemplateFromFilename(docxFilename string) (*docxTemplate, error) {
	docxBytes, err := os.ReadFile(docxFilename)
	if err != nil {
		fmt.Println("Error reading DOCX file:", err)
		return nil, err
	}

	bytesReader := bytes.NewReader(docxBytes)
	if bytesReader == nil {
		return nil, fmt.Errorf("unable to get bytes reader for DOCX file %s",
			docxFilename)
	}

	docxReader, err := zip.NewReader(bytesReader, int64(len(docxBytes)))
	if err != nil {
		return nil, fmt.Errorf("unable to get zip reader for DOCX file %s: %v", docxFilename, err)
	}

	return &docxTemplate{
		outputFilename: docxFilename,
		bytes:          docxBytes,
		reader:         docxReader,
		output:         bytes.Buffer{},
		media:          []docx.Media{},
		rel:            &docx.Relationship{},
		relMedia:       []docx.MediaRel{},
	}, nil
}

func (dt *docxTemplate) Media(filename string, data []byte) {
	dt.media = append(dt.media, docx.Media{
		Filename: filename,
		Data:     data,
	})
}

func (dt *docxTemplate) Apply(templateValues any) error {
	zipWriter := zip.NewWriter(&dt.output)

	for _, m := range dt.media {
		filepath := path.Join("word/media", m.Filename)
		err := utils.WriteFile(filepath, zipWriter, m.Data)
		if err != nil {
			return fmt.Errorf("unable to write media file %s: %w", filepath, err)
		}
	}

	ctFile := "[Content_Types].xml"
	relFile := "word/_rels/document.xml.rels"
	for _, f := range dt.reader.File {
		if f.Name != relFile {
			continue
		}

		relData, err := utils.ReadZipFileContent(f)
		if err != nil {
			fmt.Println("unable to read rel file")
			break
		}

		dt.rel, err = docx.ParseRelationship(relData)
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
	for _, f := range dt.reader.File {
		if slices.Contains(toSkip, f.Name) {
			continue
		}

		matchedChartsRel := chartsRelMatcher.MatchString(f.Name)
		if matchedChartsRel {
			fileContent, err := utils.ReadZipFileContent(f)
			if err != nil {
				return fmt.Errorf("unable to read chart rel file %s: %w", f.Name, err)
			}

			chartsRelationships, _ := docx.ParseRelationship(fileContent)
			for _, relationship := range chartsRelationships.Relationships {
				if !xlsxMatcher.MatchString(relationship.Target) {
					continue
				}

				formattedTarget := strings.Replace(relationship.Target, "../", "word/", 1)
				chartFilename, err := utils.ExtractChartName(f.Name)
				if err != nil {
					return fmt.Errorf("unable to extract chart name from file %s: %w", f.Name, err)
				}
				chartRelToTargetXlsx[chartFilename] = formattedTarget
			}
		}

		matchedChart := chartsMatcher.MatchString(f.Name)
		if matchedChart {
			continue
		}

		matchedXlsx := xlsxMatcher.MatchString(f.Name)
		if matchedXlsx {
			err := xlsx.WriteXLSXIntoZip(f, zipWriter, templateValues)
			if err != nil {
				return fmt.Errorf("unable to write XLSX file %s: %w", f.Name, err)
			}
			continue
		}

		// I don't know how many headers/footers there are, so I use a regex
		matchedHeaderFooterDocument := headerFooterDocumentMatcher.MatchString(f.Name)
		if !matchedHeaderFooterDocument {
			err := utils.CopyOriginalFile(f, zipWriter)
			if err != nil {
				return fmt.Errorf("unable to copy original file %s: %w", f.Name, err)
			}
			continue
		}

		media, err := docx.ApplyTemplate(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to file %s: %w", f.Name, err)
		}

		dt.relMedia = append(dt.relMedia, media...)
	}

	for _, f := range dt.reader.File {
		matchedChart := chartsMatcher.MatchString(f.Name)
		if !matchedChart {
			continue
		}

		fileContent, err := utils.ReadZipFileContent(f)
		if err != nil {
			return fmt.Errorf("unable to read chart file %s: %w", f.Name, err)
		}

		fileContent, err = xlsx.ApplyTemplateToChart(f, templateValues, fileContent)
		if err != nil {
			return fmt.Errorf("unable to apply template to file %s: %w", f.Name, err)
		}

		chartFileNumber, err := utils.ExtractChartName(f.Name)
		if err != nil {
			return fmt.Errorf("unable to extract chart name from file %s: %w", f.Name, err)
		}

		xlsxFileTarget := chartRelToTargetXlsx[chartFileNumber]
		fileContent, err = xlsx.UpdateChart(fileContent, xlsx.XlsxFiles[xlsxFileTarget].ChartNumbers)
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

	if len(dt.relMedia) != 0 {
		dt.rel.AddMediaToRels(dt.relMedia)
		for _, f := range dt.reader.File {
			if f.Name != relFile {
				continue
			}

			xmlContent, err := dt.rel.ToXML()
			if err != nil {
				fmt.Println("unable to marshal rels:", err)
				break
			}

			err = utils.ReplaceFileContent(f, zipWriter, []byte(xmlContent))
			if err != nil {
				fmt.Println("unable to replace rel file:", err)
				break
			}
		}
		for _, f := range dt.reader.File {
			if f.Name != ctFile {
				continue
			}

			ctData, err := utils.ReadZipFileContent(f)
			if err != nil {
				fmt.Println("unable to read content types file")
				break
			}

			contentTypes, err := docx.ParseContentTypes(ctData)
			if err != nil {
				fmt.Println("unable to parse content types file:", err)
			}

			contentTypes.EnsureImageDefaults("png", "image/png")
			updatedCt, err := contentTypes.ToXML()
			if err != nil {
				fmt.Println("unable to marshal rels:", err)
				break
			}

			err = utils.ReplaceFileContent(f, zipWriter, []byte(updatedCt))
			if err != nil {
				fmt.Println("unable to replace rel file:", err)
				break
			}
			break
		}
	}

	err := zipWriter.Close()
	if err != nil {
		return fmt.Errorf("unable to close zip writer: %w", err)
	}

	return nil
}

// Save iterates over filenames and write the output docx for the first non esistent file.
// If a single filename string is provided, the file gets overwritten.
// If no filenames are provided, it saves the file with a timestamp or the provided original filename
// if the docxTemplate object was created with the NewDocxTemplateFromFilename function.
func (dt *docxTemplate) Save(filenames ...string) error {
	filename := fmt.Sprintf("output_%s", dt.outputFilename)
	if dt.outputFilename == "" {
		filename = fmt.Sprintf("output_%s.docx", time.Now().Format("20060102150405"))
	}

	if len(filenames) == 1 {
		filename = filenames[0]
	}
	if len(filenames) > 1 {
		var err error
		filename, err = file.FindFirstMissingFile(filenames)
		if err != nil {
			fmt.Printf("The filenames provided seems to be already used, saving on '%s'\n", filename)
		}
	}

	return os.WriteFile(filename, dt.output.Bytes(), 0644)
}

func (dt *docxTemplate) Bytes() []byte {
	return dt.output.Bytes()
}
