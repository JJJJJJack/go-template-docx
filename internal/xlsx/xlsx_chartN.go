package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"text/template"

	"github.com/JJJJJJack/go-template-docx/internal/docx"
	"github.com/JJJJJJack/go-template-docx/internal/utils"
)

type V struct {
	XMLName xml.Name `xml:"v"`
	Value   string   `xml:",chardata"`
}

type Pt struct {
	V V `xml:"v"`
}

type NumCache struct {
	Pts []Pt `xml:"pt"`
}

type NumRef struct {
	NumCache NumCache `xml:"numCache"`
}

type Val struct {
	NumRef NumRef `xml:"numRef"`
}

type Ser struct {
	Val Val `xml:"val"`
}

type BarChart struct {
	Series []Ser `xml:"ser"`
}

type PlotArea struct {
	BarChart BarChart `xml:"barChart"`
}

type Chart struct {
	PlotArea PlotArea `xml:"plotArea"`
}

type ChartSpace struct {
	Chart Chart `xml:"chart"`
}

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

func ApplyTemplateToXml(f *zip.File, templateValues any) ([]byte, error) {
	fileContent, err := utils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read chart file %s: %w", f.Name, err)
	}

	tmpl, err := template.New(f.Name).
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

// TODO: switch to xml parsing
func ReplaceSharedStringIndicesWithValues(fileContent []byte, values map[int]string) ([]byte, []string, error) {
	re := regexp.MustCompile(`<c[^>]*t="s"[^>]*>.*?<v>(\d+)</v>.*?</c>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	valuesOrderedByAppearance := []string{}
	for _, match := range matches {
		n, err := strconv.Atoi(match[1])
		if err != nil {
			return nil, nil, fmt.Errorf("unable to convert index %s to int: %w", match[1], err)
		}

		if value, ok := values[n]; ok {
			removedRefToSharedString := strings.Replace(match[0], `t="s"`, "", 1)

			oldV := fmt.Sprintf("<v>%d</v>", n)
			newV := fmt.Sprintf("<v>%s</v>", value)

			replace := strings.Replace(removedRefToSharedString, oldV, newV, 1)

			fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(replace), 1)

			valuesOrderedByAppearance = append(valuesOrderedByAppearance, value)
		}
	}

	return fileContent, valuesOrderedByAppearance, nil
}
