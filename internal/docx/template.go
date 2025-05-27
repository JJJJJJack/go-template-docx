package docx

import (
	"archive/zip"
	"bytes"
	"io"
	"regexp"
	"strings"
	"text/template"
)

func patchXML(srcXML string) string {
	// Fix separated [[
	re := regexp.MustCompile(`\[([\s\S]*?)\[`)
	srcXML = re.ReplaceAllString(srcXML, "[[")

	// Fix separated ]]
	re = regexp.MustCompile(`\]([\s\S]*?)\]`)
	srcXML = re.ReplaceAllString(srcXML, "]]")

	// Remove unecessary XML tags
	re = regexp.MustCompile(`\[\[[\s\S]*?\]\]`)
	matches := re.FindAllString(srcXML, -1)
	for _, match := range matches {
		xmlRegex := regexp.MustCompile(`(<\s*\/?[\w-:.]+(\s+[^>]*?)?\s*>)`)
		templateText := xmlRegex.ReplaceAllString(match, "")
		srcXML = strings.ReplaceAll(srcXML, match, templateText)
	}

	return srcXML
}

func applyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]mediaRel, error) {
	documentFile, err := f.Open()
	if err != nil {
		return nil, err
	}

	documentXML, err := io.ReadAll(documentFile)
	if err != nil {
		return nil, err
	}
	documentXML = []byte(patchXML(string(documentXML)))

	tmpl, err := template.New("report-template").
		Delims("[[", "]]").
		Funcs(template.FuncMap{
			"toImage": toImage,
		}).
		Parse(string(documentXML))
	if err != nil {
		return nil, err
	}

	appliedTemplate := bytes.Buffer{}
	tmpl.Execute(&appliedTemplate, data)

	output, media, err := applyImages(appliedTemplate.String())
	if err != nil {
		return nil, err
	}
	output = postProcessing(output)

	newDocumentXMLFile, _ := zipWriter.CreateHeader(&zip.FileHeader{
		Name:   f.Name,
		Method: f.FileHeader.Method,
	})
	newDocumentXMLFile.Write([]byte(output))

	return media, nil
}
