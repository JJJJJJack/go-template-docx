package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"strings"
	"text/template"
)

func patchXML(srcXML string) string {
	// Fix separated {{
	re := regexp.MustCompile(`\{([^\}]*?)\{`)
	srcXML = re.ReplaceAllString(srcXML, "{{")

	// Fix separated }}
	re = regexp.MustCompile(`\}([^\{]*?)\}`)
	srcXML = re.ReplaceAllString(srcXML, "}}")

	// Remove unecessary XML tags
	re = regexp.MustCompile(`\{\{[\s\S]*?\}\}`)
	matches := re.FindAllString(srcXML, -1)
	for _, match := range matches {
		xmlRegex := regexp.MustCompile(`(<\s*\/?[\w-:.]+(\s+[^>]*?)?[\s\/]*>)`)
		templateText := xmlRegex.ReplaceAllString(match, "")
		srcXML = strings.ReplaceAll(srcXML, match, templateText)
	}

	return srcXML
}

func applyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]mediaRel, error) {
	documentFile, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("unable to open document file %s: %w", f.Name, err)
	}

	documentXML, err := io.ReadAll(documentFile)
	if err != nil {
		return nil, fmt.Errorf("unable to read document file %s: %w", f.Name, err)
	}
	documentXML = []byte(patchXML(string(documentXML)))

	tmpl, err := template.New("report-template").
		Funcs(template.FuncMap{
			"toImage": toImage,
		}).
		Parse(string(documentXML))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template in file %s: %w", f.Name, err)
	}

	appliedTemplate := bytes.Buffer{}
	err = tmpl.Execute(&appliedTemplate, data)
	if err != nil {
		return nil, fmt.Errorf("unable to execute template in file %s: %w", f.Name, err)
	}

	output, media, err := applyImages(appliedTemplate.String())
	if err != nil {
		return nil, fmt.Errorf("unable to apply images in file %s: %w", f.Name, err)
	}
	output = postProcessing(output)

	newDocumentXMLFile, _ := zipWriter.CreateHeader(&zip.FileHeader{
		Name:   f.Name,
		Method: f.FileHeader.Method,
	})
	newDocumentXMLFile.Write([]byte(output))

	return media, nil
}
