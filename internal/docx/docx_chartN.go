package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"text/template"
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
func UpdateDocxChartWithXlsxChartValues(fileContent []byte, values []string) ([]byte, error) {
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

func ApplyTemplateToChart(f *zip.File, templateValues any, fileContent []byte) ([]byte, error) {
	tmpl, err := template.New(f.Name).
		Parse(patchXML(string(fileContent)))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template: %w", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, templateValues); err != nil {
		return nil, fmt.Errorf("unable to execute template: %w", err)
	}

	return buf.Bytes(), nil
}
