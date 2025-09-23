package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"text/template"

	"github.com/JJJJJJack/go-template-docx/internal/docx"
)

// ApplyTemplateToCells applies the templateValues to the given file content and returns the modified content.
func ApplyTemplateToCells(f *zip.File, templateValues any, fileContent []byte) ([]byte, error) {
    tmpl, err := template.New(f.Name).
        Option("missingkey=error").
        Funcs(template.FuncMap{
            "toNumberCell": ToNumberCell,
        }).
        Parse(docx.PatchXml(string(fileContent)))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, templateValues); err != nil {
		return nil, fmt.Errorf("unable to execute template: %w", err)
	}

	return buf.Bytes(), nil
}
