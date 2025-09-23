package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"text/template"

	goziputils "github.com/JJJJJJack/go-zip-utils"
)

// TODO: parse and unmarshal xml instead of using regex
func UpdateChart(fileContent []byte, cellAndValues map[string]string) ([]byte, error) {
	// the blockRe captures both strCache and numCache because it would otherwise match the first <c:f>...</c:f> with the last <c:numCache>...</c:numCache>
	blockRe := regexp.MustCompile(`(?s)<c:f>(Sheet\d+!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+))</c:f>.*?<c:(?:strCache|numCache)>(.*?)</c:(?:strCache|numCache)>`)
	ptRe := regexp.MustCompile(`(?s)<c:pt idx="(\d+)">.*?<c:v>(.*?)</c:v>.*?</c:pt>`)

	updated := blockRe.ReplaceAllFunc(fileContent, func(block []byte) []byte {
		m := blockRe.FindSubmatch(block)
		if len(m) < 6 {
			return block
		}

		col := string(m[2])                       // "A"
		startRow, _ := strconv.Atoi(string(m[3])) // 2
		cache := string(m[5])                     // contents of <c:strCache> or <c:numCache>

		// Iterate over <c:pt>
		cacheUpdated := ptRe.ReplaceAllStringFunc(cache, func(pt string) string {
			pm := ptRe.FindStringSubmatch(pt)
			if len(pm) < 3 {
				return pt
			}

			idx, _ := strconv.Atoi(pm[1]) // idx=0,1,2,3...
			cell := fmt.Sprintf("%s%d", col, startRow+idx)

			if number, ok := cellAndValues[cell]; ok {
				oldVal := fmt.Sprintf("<c:v>%s</c:v>", pm[2])
				newVal := fmt.Sprintf("<c:v>%s</c:v>", number)
				return strings.Replace(pt, oldVal, newVal, 1)
			}
			return pt
		})

		// Put updated cache back
		return []byte(strings.Replace(string(block), cache, cacheUpdated, 1))
	})

	return updated, nil
}

func ApplyTemplateToXml(f *zip.File, templateValues any, templateFuncs template.FuncMap) ([]byte, error) {
	fileContent, err := goziputils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read chart file '%s': %w", f.Name, err)
	}

	tmpl, err := template.New(f.Name).
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
