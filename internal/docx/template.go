package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"strings"
	"text/template"

	"github.com/JJJJJJack/go-template-docx/internal/utils"
)

func PatchXML(srcXML string) string {
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

func ApplyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]MediaRel, error) {
	documentXML, err := utils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read document file %s: %w", f.Name, err)
	}
	documentXML = []byte(PatchXML(string(documentXML)))

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
