package gotemplatedocx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path"
	"regexp"
	"strings"
	"time"

	"github.com/JJJJJJack/go-template-docx/internal/docx"
	"github.com/JJJJJJack/go-template-docx/internal/file"
	"github.com/JJJJJJack/go-template-docx/internal/utils"
	"github.com/JJJJJJack/go-template-docx/internal/xlsx"
)

type DocxTemplate struct {
	outputFilename string
	bytes          []byte
	reader         *zip.Reader
	output         bytes.Buffer
	rel            *docx.Relationship
	relMedia       []docx.MediaRel
	media          []docx.Media
}

// NewDocxTemplateFromBytes creates a new DocxTemplate object from the provided DOCX file bytes.
// The DocxTemplate object can be used through the exposed high-level APIs.
func NewDocxTemplateFromBytes(docxBytes []byte) (*DocxTemplate, error) {
	bytesReader := bytes.NewReader(docxBytes)
	if bytesReader == nil {
		return nil, fmt.Errorf("unable to create bytes reader for DOCX file")
	}

	docxReader, err := zip.NewReader(bytesReader, int64(len(docxBytes)))
	if err != nil {
		return nil, fmt.Errorf("unable to create zip reader for DOCX file: %w", err)
	}

	return &DocxTemplate{
		outputFilename: "",
		bytes:          docxBytes,
		reader:         docxReader,
		output:         bytes.Buffer{},
		media:          []docx.Media{},
		rel:            &docx.Relationship{},
		relMedia:       []docx.MediaRel{},
	}, nil
}

// NewDocxTemplateFromFilename creates a new DocxTemplate object from the provided DOCX filename (reading from disk).
// The DocxTemplate object can be used through the exposed high-level APIs.
func NewDocxTemplateFromFilename(docxFilename string) (*DocxTemplate, error) {
	docxBytes, err := os.ReadFile(docxFilename)
	if err != nil {
		return nil, fmt.Errorf("unable to read file %s: %w", docxFilename, err)
	}

	bytesReader := bytes.NewReader(docxBytes)
	if bytesReader == nil {
		return nil, fmt.Errorf("unable to create bytes reader for DOCX file %s",
			docxFilename)
	}

	docxReader, err := zip.NewReader(bytesReader, int64(len(docxBytes)))
	if err != nil {
		return nil, fmt.Errorf("unable to create zip reader for DOCX file %s: %w", docxFilename, err)
	}

	return &DocxTemplate{
		outputFilename: docxFilename,
		bytes:          docxBytes,
		reader:         docxReader,
		output:         bytes.Buffer{},
		media:          []docx.Media{},
		rel:            &docx.Relationship{},
		relMedia:       []docx.MediaRel{},
	}, nil
}

// Media adds a media file to the DocxTemplate object.
// Supported media types are currently limited to JPEG and PNG images.
// The filename match the string you pass in the template expression using the "toImage" function.
// as example {{ toImage "computer.png" }} will load the docx.Media that have "computer.png" as its filename.
// The data should be the byte content of the media file.
func (dt *DocxTemplate) Media(filename string, data []byte) {
	dt.media = append(dt.media, docx.Media{
		Filename: filename,
		Data:     data,
	})
}

// Apply applies the template with the provided values to the DOCX file.
// The templateValues parameter can be any type that can be marshalled to JSON.
func (dt *DocxTemplate) Apply(templateValues any) error {
	zipWriter := zip.NewWriter(&dt.output)

	zipMap := make(utils.ZipMap)
	for _, f := range dt.reader.File {
		zipMap[f.Name] = f
	}

	document, err := docx.ParseDocumentMeta(zipMap)
	if err != nil {
		return fmt.Errorf("unable to parse document metadata: %w", err)
	}

	// Copy all files except the ones that will be processed
	documentRelsFilename := "word/_rels/document.xml.rels"
	contentTypesFilename := "[Content_Types].xml"
	chartsMatcher := regexp.MustCompile(`word/charts/chart\d*?\.xml`)
	headerFooterDocumentMatcher := regexp.MustCompile(`word/(header|footer|document)\d*?\.xml`)
	xlsxMatcher := regexp.MustCompile(`/embeddings/Microsoft_Excel_Worksheet\d*?\.xlsx`)
	for _, f := range dt.reader.File {
		switch {
		case
			f.Name == documentRelsFilename,
			f.Name == contentTypesFilename,
			chartsMatcher.MatchString(f.Name),
			xlsxMatcher.MatchString(f.Name),
			headerFooterDocumentMatcher.MatchString(f.Name):
			continue
		}

		err := utils.CopyOriginalFile(f, zipWriter)
		if err != nil {
			return fmt.Errorf("unable to copy original file '%s': %w", f.Name, err)
		}
	}

	// Edit [Content_Types].xml if media files are provided
	ctFile := zipMap["[Content_Types].xml"]
	ctData, err := utils.ReadZipFileContent(ctFile)
	if err != nil {
		return fmt.Errorf("unable to read content types file '%s': %w", ctFile.Name, err)
	}

	contentTypes, err := docx.ParseContentTypes(ctData)
	if err != nil {
		return fmt.Errorf("unable to parse content types file '%s': %w", ctFile.Name, err)
	}

	for _, media := range dt.media {
		ext := path.Ext(media.Filename)

		switch strings.ToLower(ext) {
		case ".jpg", ".jpeg", "jfif":
			contentTypes.AddDefaultUnique("jpeg", "image/jpeg")
		case ".png":
			contentTypes.AddDefaultUnique("png", "image/png")
		default:
			fmt.Println("Unsupported media file type (only accepting jpg/png for now):", media.Filename)
			continue
		}
	}

	updatedCt, err := contentTypes.ToXml()
	if err != nil {
		return fmt.Errorf("unable to marshal content types to XML: %w", err)
	}

	err = utils.ReplaceFileContent(ctFile, zipWriter, []byte(updatedCt))
	if err != nil {
		return fmt.Errorf("unable to replace content types file '%s': %w", ctFile.Name, err)
	}

	// Put loaded medias into the new docx file
	for _, m := range dt.media {
		filepath := path.Join("word/media", m.Filename)
		err := utils.ZipWriteFile(filepath, zipWriter, m.Data)
		if err != nil {
			return fmt.Errorf("unable to write media file '%s': %w", filepath, err)
		}
	}

	relData, err := utils.ReadZipFileContent(zipMap[documentRelsFilename])
	if err != nil {
		return fmt.Errorf("unable to read rel file '%s': %w", documentRelsFilename, err)
	}

	dt.rel, err = docx.ParseRelationship(relData)
	if err != nil {
		return fmt.Errorf("unable to parse rel file '%s': %w", documentRelsFilename, err)
	}

	// Map chart files to their target XLSX files
	chartRelToTargetXlsx := make(map[string]string)
	for i := 1; ; i++ {
		relsChartFilename := fmt.Sprintf("word/charts/_rels/chart%d.xml.rels", i)
		f := zipMap[relsChartFilename]
		if f == nil {
			break
		}

		fileContent, err := utils.ReadZipFileContent(f)
		if err != nil {
			return fmt.Errorf("unable to read chart rel file '%s': %w", f.Name, err)
		}

		chartsRelationships, _ := docx.ParseRelationship(fileContent)
		for _, relationship := range chartsRelationships.Relationships {
			if !xlsxMatcher.MatchString(relationship.Target) {
				continue
			}

			targetXlsxFilename := strings.Replace(relationship.Target, "../", "word/", 1)
			chartFilename, err := utils.ExtractChartFilename(f.Name)
			if err != nil {
				return fmt.Errorf("unable to extract chart name from file '%s': %w", f.Name, err)
			}
			chartRelToTargetXlsx[chartFilename] = targetXlsxFilename
		}
	}

	// Apply template to the XLSX files
	for i := 0; ; i++ {
		xlsxFilename := fmt.Sprintf("word/embeddings/Microsoft_Excel_Worksheet%d.xlsx", i)
		if i == 0 {
			xlsxFilename = "word/embeddings/Microsoft_Excel_Worksheet.xlsx"
		}
		f := zipMap[xlsxFilename]
		if f == nil {
			break
		}

		err := xlsx.WriteXlsxIntoZip(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to XLSX file '%s': %w", f.Name, err)
		}
	}

	// Apply template to the header files
	for i := 1; ; i++ {
		headerFilename := fmt.Sprintf("word/header%d.xml", i)
		f := zipMap[headerFilename]
		if f == nil {
			break
		}

		media, err := document.ApplyTemplate(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to header file '%s': %w", f.Name, err)
		}

		dt.relMedia = append(dt.relMedia, media...)
	}

	// Apply template to the footer files
	for i := 1; ; i++ {
		footerFilename := fmt.Sprintf("word/footer%d.xml", i)
		f := zipMap[footerFilename]
		if f == nil {
			break
		}

		media, err := document.ApplyTemplate(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to footer file '%s': %w", f.Name, err)
		}

		dt.relMedia = append(dt.relMedia, media...)
	}

	// Apply template to the main document file
	documentFile := zipMap["word/document.xml"]
	if documentFile == nil {
		return fmt.Errorf("word/document.xml not found in the DOCX file")
	}

	media, err := document.ApplyTemplate(documentFile, zipWriter, templateValues)
	if err != nil {
		return fmt.Errorf("unable to apply template to document file: %w", err)
	}

	dt.relMedia = append(dt.relMedia, media...)

	// Apply template to the chart files
	for i := 1; ; i++ {
		chartN := fmt.Sprintf("word/charts/chart%d.xml", i)
		f := zipMap[chartN]
		if f == nil {
			break
		}

		fileContent, err := xlsx.ApplyTemplateToXml(f, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to chart file '%s': %w", f.Name, err)
		}

		chartFilename, err := utils.ExtractChartFilename(f.Name)
		if err != nil {
			return fmt.Errorf("unable to extract chart name from file '%s': %w", f.Name, err)
		}

		xlsxFileTarget := chartRelToTargetXlsx[chartFilename]
		fileContent, err = xlsx.UpdateChart(fileContent, xlsx.XlsxFiles[xlsxFileTarget].ChartNumbers)
		if err != nil {
			return fmt.Errorf("unable to update preview chart file '%s': %w", f.Name, err)
		}

		err = utils.RewriteFileIntoZipWriter(f, zipWriter, fileContent)
		if err != nil {
			return fmt.Errorf("unable to rewrite chart file '%s': %w", f.Name, err)
		}
	}

	if len(dt.relMedia) != 0 {
		dt.rel.AddMediaToRels(dt.relMedia)

		documentRelFile := zipMap[documentRelsFilename]

		xmlContent, err := dt.rel.ToXML()
		if err != nil {
			return fmt.Errorf("unable to marshal rels: %w", err)
		}

		err = utils.ReplaceFileContent(documentRelFile, zipWriter, []byte(xmlContent))
		if err != nil {
			return fmt.Errorf("unable to replace rel file '%s': %w", documentRelsFilename, err)
		}
	}

	err = zipWriter.Close()
	if err != nil {
		return fmt.Errorf("unable to close zip writer: %w", err)
	}

	return nil
}

// Save iterates over filenames and write the output docx for the first non existent file.
// If a single filename string is provided, the file gets overwritten.
// If no filenames are provided, it saves the file with a timestamp or the provided original filename
// if the DocxTemplate object was created with the NewDocxTemplateFromFilename function.
func (dt *DocxTemplate) Save(filenames ...string) error {
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

// Bytes returns the output bytes of the output DOCX file bytes
// (empty if Apply was not used).
func (dt *DocxTemplate) Bytes() []byte {
	return dt.output.Bytes()
}
