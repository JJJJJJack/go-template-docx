package gotemplatedocx

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"
	"text/template"

	"github.com/JJJJJJack/go-template-docx/internal/docx"
	"github.com/JJJJJJack/go-template-docx/xml"
	goziputils "github.com/JJJJJJack/go-zip-utils"
)

type docxTemplate struct {
	bytes    []byte
	reader   *zip.Reader
	output   bytes.Buffer
	rel      *docx.Relationship
	relMedia []docx.MediaRel
	// filename : data
	media               docx.MediaMap
	xlsxChartsMeta      xlsxChartsMap
	templateFuncs       template.FuncMap
	filesPreProcessors  xml.HandlersMap
	filesPostProcessors xml.HandlersMap
}

// NewDocxTemplateFromBytes creates a new docxTemplate object from the provided DOCX file bytes.
// The docxTemplate object can be used through the exposed high-level APIs.
func NewDocxTemplateFromBytes(docxBytes []byte) (*docxTemplate, error) {
	bytesReader := bytes.NewReader(docxBytes)
	if bytesReader == nil {
		return nil, fmt.Errorf("unable to create bytes reader for DOCX file")
	}

	docxReader, err := zip.NewReader(bytesReader, int64(len(docxBytes)))
	if err != nil {
		return nil, fmt.Errorf("unable to create zip reader for DOCX file: %w", err)
	}

	return &docxTemplate{
		bytes:               docxBytes,
		reader:              docxReader,
		output:              bytes.Buffer{},
		media:               docx.MediaMap{},
		rel:                 &docx.Relationship{},
		relMedia:            []docx.MediaRel{},
		xlsxChartsMeta:      make(xlsxChartsMap),
		templateFuncs:       make(template.FuncMap),
		filesPreProcessors:  make(xml.HandlersMap),
		filesPostProcessors: make(xml.HandlersMap),
	}, nil
}

// NewDocxTemplateFromFilename creates a new docxTemplate object from the provided DOCX filename (reading from disk).
// The docxTemplate object can be used through the exposed high-level APIs.
func NewDocxTemplateFromFilename(docxFilename string) (*docxTemplate, error) {
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

	return &docxTemplate{
		bytes:               docxBytes,
		reader:              docxReader,
		output:              bytes.Buffer{},
		media:               make(docx.MediaMap),
		rel:                 &docx.Relationship{},
		relMedia:            []docx.MediaRel{},
		xlsxChartsMeta:      make(xlsxChartsMap),
		templateFuncs:       make(template.FuncMap),
		filesPreProcessors:  make(xml.HandlersMap),
		filesPostProcessors: make(xml.HandlersMap),
	}, nil
}

// Media adds a media file to the docxTemplate object.
// Supported media types are currently limited to JPEG and PNG images.
// The filename match the string you pass in the template expression using the image function.
// For example {{ image "computer.png" }} will load the docx.Media that have "computer.png" as its filename.
// The data should be the byte content of the media file.
func (dt *docxTemplate) Media(filename string, data []byte) {
	filename = filepath.Base(filename)

	dt.media[filename] = data
}

// AddTemplateFuncs adds your custom template functions to evaluate when applying the template.
// Existing functions will be shadowed if the same name is used.
func (dt *docxTemplate) AddTemplateFuncs(funcMap template.FuncMap) {
	dt.templateFuncs = funcMap
}

// AddPreProcessors adds XML pre-processing map in which the key is the XML file path
// (e.g., "word/document.xml") and the value is a list of functions to be applied to that file
// before the template has been applied.
func (dt *docxTemplate) AddPreProcessors(filesPreProcessors xml.HandlersMap) {
	dt.filesPreProcessors = filesPreProcessors
}

// AddPostProcessors adds XML post-processing map in which the key is the XML file path
// (e.g., "word/document.xml") and the value is a list of functions to be applied to that file
// after the template has been applied.
func (dt *docxTemplate) AddPostProcessors(filesPostProcessors xml.HandlersMap) {
	dt.filesPostProcessors = filesPostProcessors
}

// Apply applies the template with the provided values to the DOCX file.
// The templateValues parameter can be any type that can be marshalled to JSON.
func (dt *docxTemplate) Apply(templateValues any) error {
	switch v := templateValues.(type) {
	case []byte:
		err := json.Unmarshal(v, &templateValues)
		if err != nil {
			return fmt.Errorf("error unmarshalling templateValues: %w", err)
		}
	}

	// custom user pre processing
	if len(dt.filesPreProcessors) > 0 {
		preZipMap, err := goziputils.NewZipMapFromBytes(dt.bytes)
		if err != nil {
			return fmt.Errorf("unable to create DOCX zip map for pre-processing: %w", err)
		}

		preOutput := bytes.Buffer{}

		preZipWriter := zip.NewWriter(&preOutput)

		for filename, f := range preZipMap {
			processors := dt.filesPreProcessors[filename]
			if len(processors) == 0 {
				err := goziputils.CopyFile(preZipWriter, f)
				if err != nil {
					return fmt.Errorf("unable to copy original file '%s' during pre-processing: %w", f.Name, err)
				}

				continue
			}

			fileContent, err := goziputils.ReadZipFileContent(f)
			if err != nil {
				return fmt.Errorf("unable to read file '%s' for pre-processing: %w", f.Name, err)
			}

			preOutput := string(fileContent)
			for _, processor := range processors {
				preOutput, err = processor(preOutput)
				if err != nil {
					return fmt.Errorf("error pre processing file '%s': %w", f.Name, err)
				}
			}

			err = goziputils.RewriteFileIntoZipWriter(preZipWriter, f, []byte(preOutput))
			if err != nil {
				return fmt.Errorf("unable to rewrite pre-processed file '%s': %w", f.Name, err)
			}
		}

		err = preZipWriter.Close()
		if err != nil {
			return fmt.Errorf("unable to close zip writer after pre-processing: %w", err)
		}

		dt.bytes = preOutput.Bytes()
	}

	zipWriter := zip.NewWriter(&dt.output)

	docxZipMap, err := goziputils.NewZipMapFromBytes(dt.bytes)
	if err != nil {
		return fmt.Errorf("unable to create DOCX zip map: %w", err)
	}

	document, err := docx.ParseDocumentMeta(docxZipMap, dt.templateFuncs)
	if err != nil {
		return fmt.Errorf("unable to parse document metadata: %w", err)
	}

	// Copy all files except the ones that will be processed
	documentRelsFilename := "word/_rels/document.xml.rels"
	contentTypesFilename := "[Content_Types].xml"
	chartsMatcher := regexp.MustCompile(`word/charts/chart\d*?\.xml`)
	xlsxMatcher := regexp.MustCompile(`/embeddings/Microsoft_Excel_Worksheet\d*?\.xlsx`)
	headerFooterDocumentMatcher := regexp.MustCompile(`word/(header|footer|document)\d*?\.xml`)
	for filename, f := range docxZipMap {
		switch {
		case
			filename == documentRelsFilename,
			filename == contentTypesFilename,
			chartsMatcher.MatchString(filename),
			xlsxMatcher.MatchString(filename),
			headerFooterDocumentMatcher.MatchString(filename):
			continue
		}

		err := goziputils.CopyFile(zipWriter, f)
		if err != nil {
			return fmt.Errorf("unable to copy original file '%s': %w", f.Name, err)
		}
	}

	// Edit [Content_Types].xml if media files are provided
	ctFile := docxZipMap[contentTypesFilename]
	ctData, err := goziputils.ReadZipFileContent(ctFile)
	if err != nil {
		return fmt.Errorf("unable to read content types file '%s': %w", ctFile.Name, err)
	}

	contentTypes, err := docx.ParseContentTypes(ctData)
	if err != nil {
		return fmt.Errorf("unable to parse content types file '%s': %w", ctFile.Name, err)
	}

	for filename := range dt.media {
		ext := path.Ext(filename)

		switch strings.ToLower(ext) {
		case ".jpg", ".jpeg", "jfif":
			contentTypes.AddDefaultUnique("jpeg", "image/jpeg")
		case ".png":
			contentTypes.AddDefaultUnique("png", "image/png")
		default:
			fmt.Println("Unsupported media file type (only accepting jpg/png for now):", filename)
			continue
		}
	}

	updatedCt, err := contentTypes.ToXml()
	if err != nil {
		return fmt.Errorf("unable to marshal content types to XML: %w", err)
	}

	err = goziputils.RewriteFileIntoZipWriter(zipWriter, ctFile, []byte(updatedCt))
	if err != nil {
		return fmt.Errorf("unable to replace content types file '%s': %w", ctFile.Name, err)
	}

	// Put loaded medias into the new docx file
	for filename, data := range dt.media {
		filepath := path.Join("word/media", filename)
		err := goziputils.WriteFile(zipWriter, filepath, data)
		if err != nil {
			return fmt.Errorf("unable to write media file '%s': %w", filepath, err)
		}
	}

	relData, err := goziputils.ReadZipFileContent(docxZipMap[documentRelsFilename])
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
		f := docxZipMap[relsChartFilename]
		if f == nil {
			break
		}

		fileContent, err := goziputils.ReadZipFileContent(f)
		if err != nil {
			return fmt.Errorf("unable to read chart rel file '%s': %w", f.Name, err)
		}

		chartsRelationships, _ := docx.ParseRelationship(fileContent)
		for _, relationship := range chartsRelationships.Relationships {
			if !xlsxMatcher.MatchString(relationship.Target) {
				continue
			}

			targetXlsxFilename := strings.Replace(relationship.Target, "../", "word/", 1)
			chartFilename, err := docx.ExtractChartFilename(f.Name)
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
		f := docxZipMap[xlsxFilename]
		if f == nil {
			break
		}

		err := dt.writeXlsxIntoZip(f, zipWriter, templateValues)
		if err != nil {
			return fmt.Errorf("unable to apply template to XLSX file '%s': %w", f.Name, err)
		}
	}

	// Apply template to the header files
	for i := 1; ; i++ {
		headerFilename := fmt.Sprintf("word/header%d.xml", i)
		f := docxZipMap[headerFilename]
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
		f := docxZipMap[footerFilename]
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
	documentFile := docxZipMap["word/document.xml"]
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

		f := docxZipMap[chartN]
		if f == nil {
			break
		}

		fileContent, err := docx.ApplyTemplateToXml(f, templateValues, dt.templateFuncs)
		if err != nil {
			return fmt.Errorf("unable to apply template to chart file '%s': %w", f.Name, err)
		}

		chartFilename, err := docx.ExtractChartFilename(f.Name)
		if err != nil {
			return fmt.Errorf("unable to extract chart name from file '%s': %w", f.Name, err)
		}

		xlsxFileTarget := chartRelToTargetXlsx[chartFilename]
		fileContent, err = docx.UpdateChart(fileContent, dt.xlsxChartsMeta[xlsxFileTarget].chartNumbers)
		if err != nil {
			return fmt.Errorf("unable to update preview chart file '%s': %w", f.Name, err)
		}

		err = goziputils.RewriteFileIntoZipWriter(zipWriter, f, fileContent)
		if err != nil {
			return fmt.Errorf("unable to rewrite chart file '%s': %w", f.Name, err)
		}
	}

	documentRelFile := docxZipMap[documentRelsFilename]
	documentRelContent, err := goziputils.ReadZipFileContent(documentRelFile)
	if err != nil {
		return fmt.Errorf("unable to read rel file '%s': %w", documentRelsFilename, err)
	}

	if len(dt.relMedia) != 0 {
		dt.rel.AddMediaToRels(dt.relMedia)

		documentRelContent, err = dt.rel.ToXml()
		if err != nil {
			return fmt.Errorf("unable to marshal rels: %w", err)
		}
	}

	err = goziputils.RewriteFileIntoZipWriter(zipWriter, documentRelFile, documentRelContent)
	if err != nil {
		return fmt.Errorf("unable to replace rel file '%s': %w", documentRelsFilename, err)
	}

	err = zipWriter.Close()
	if err != nil {
		return fmt.Errorf("unable to close zip writer: %w", err)
	}

	// custom user post processing
	if len(dt.filesPostProcessors) > 0 {
		outputPost := bytes.Buffer{}
		postZipWriter := zip.NewWriter(&outputPost)

		finalZipMap, err := goziputils.NewZipMapFromBytes(dt.output.Bytes())
		if err != nil {
			return fmt.Errorf("unable to create final zip map for post-processing: %w", err)
		}

		for filename, f := range finalZipMap {
			processors := dt.filesPostProcessors[filename]
			if len(processors) == 0 {
				err := goziputils.CopyFile(postZipWriter, f)
				if err != nil {
					return fmt.Errorf("unable to copy original file '%s' during post-processing: %w", f.Name, err)
				}

				continue
			}

			fileContent, err := goziputils.ReadZipFileContent(f)
			if err != nil {
				return fmt.Errorf("unable to read file '%s' for post-processing: %w", f.Name, err)
			}

			postOutput := string(fileContent)
			for _, processor := range processors {
				postOutput, err = processor(postOutput)
				if err != nil {
					return fmt.Errorf("error post processing file '%s': %w", f.Name, err)
				}
			}

			err = goziputils.RewriteFileIntoZipWriter(postZipWriter, f, []byte(postOutput))
			if err != nil {
				return fmt.Errorf("unable to rewrite post-processed file '%s': %w", f.Name, err)
			}
		}

		err = postZipWriter.Close()
		if err != nil {
			return fmt.Errorf("unable to close zip writer after post-processing: %w", err)
		}

		dt.output = outputPost
	}

	return nil
}

// Save saves the modified docx file to the specified filename.
func (dt *docxTemplate) Save(filename string) error {
	return os.WriteFile(filename, dt.output.Bytes(), 0644)
}

// Bytes returns the output bytes of the output xlsx file bytes
// (empty if Apply was not used).
func (dt *docxTemplate) Bytes() []byte {
	return dt.output.Bytes()
}
