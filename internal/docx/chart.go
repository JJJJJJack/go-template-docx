package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"text/template"

	goziputils "github.com/JJJJJJack/go-zip-utils"
)

// TODO: parse and unmarshal xml instead of using regex
func UpdateChart(fileContent []byte, values []string) ([]byte, error) {
	re := regexp.MustCompile(`<c:val>.*?<c:v>(.*?)</c:v>.*?</c:val>`)
	matches := re.FindAllSubmatch(fileContent, -1)

	if len(matches) == 0 {
		return fileContent, nil
	}

	for _, value := range values {
		fileContent = bytes.Replace(fileContent, []byte("<c:v>0</c:v>"), []byte(fmt.Sprintf("<c:v>%s</c:v>", value)), 1)
	}

	return fileContent, nil
}

func ApplyTemplateToXml(f *zip.File, templateValues any, templateFuncs template.FuncMap) ([]byte, error) {
	fileContent, err := goziputils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read chart file '%s': %w", f.Name, err)
	}

    tmpl, err := template.New(f.Name).
        Option("missingkey=error").
        Funcs(templateFuncs).
        Parse(PatchXml(string(fileContent)))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, templateValues); err != nil {
		return nil, fmt.Errorf("unable to execute template: %w", err)
	}

	return buf.Bytes(), nil
}

// ExtractChartFilename now only works with a single submatch
func ExtractChartFilename(path string) (string, error) {
	re := regexp.MustCompile(`(chart\d+)\.xml`)
	matches := re.FindStringSubmatch(path)
	if len(matches) < 2 {
		return "", fmt.Errorf("no chart name found")
	}
	return matches[1], nil
}
