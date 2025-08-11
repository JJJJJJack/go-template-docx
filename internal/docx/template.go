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

// PatchXml removes automatically insert content between template expressions
// (EG: "{{ .Text }}" could have correctors highlights tags separating the expressions tokens).
func PatchXml(srcXml string) string {
	// Fix separated {{
	re := regexp.MustCompile(`\{([^\}]*?)\{`)
	srcXml = re.ReplaceAllString(srcXml, "{{")

	// Fix separated }}
	re = regexp.MustCompile(`\}([^\{]*?)\}`)
	srcXml = re.ReplaceAllString(srcXml, "}}")

	// Remove unecessary XML tags
	re = regexp.MustCompile(`\{\{[\s\S]*?\}\}`)
	matches := re.FindAllString(srcXml, -1)
	for _, match := range matches {
		xmlRegex := regexp.MustCompile(`(<\s*\/?[\w-:.]+(\s+[^>]*?)?[\s\/]*>)`)
		templateText := xmlRegex.ReplaceAllString(match, "")
		srcXml = strings.ReplaceAll(srcXml, match, templateText)
	}

	return srcXml
}

func ApplyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]MediaRel, error) {
	documentXML, err := utils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read document file %s: %w", f.Name, err)
	}
	documentXML = []byte(PatchXml(string(documentXML)))

	tmpl, err := template.New(f.Name).
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

	err = utils.RewriteFileIntoZipWriter(f, zipWriter, []byte(output))
	if err != nil {
		return nil, fmt.Errorf("unable to rewrite file %s in zip: %w", f.Name, err)
	}

	return media, nil
}
