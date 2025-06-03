package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"maps"
	"regexp"
	"slices"
	"sort"
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

func ReplacePreviewZerosWithXlsxChartValues(fileContent []byte, values map[int]string) ([]byte, error) {
	/* broken code, not used, but optimally will be used to replace bad regex */
	// var doc ChartSpace
	// err := xml.Unmarshal(fileContent, &doc)
	// if err != nil {
	// 	return nil, fmt.Errorf("unable to unmarshal XML: %w", err)
	// }

	indexes := slices.Collect(maps.Keys(values))
	sort.Ints(indexes)

	// for _, ser := range doc.Chart.PlotArea.BarChart.Series {
	// 	for i, pt := range ser.Val.NumRef.NumCache.Pts {
	// 		if i >= len(vals) {
	// 			continue
	// 		}
	// 		pt.V.Value = vals[i]
	// 		ser.Val.NumRef.NumCache.Pts[i] = pt
	// 	}
	// }

	// fileContent, err = xml.Marshal(doc)
	// if err != nil {
	// 	return nil, fmt.Errorf("unable to marshal XML: %w", err)
	// }

	// regexp.Compile(`</c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>(.*?)`)

	if len(indexes) == 0 {
		return fileContent, nil
	}

	re := regexp.MustCompile(`<c:val>.*?<c:v>(.*?)</c:v>.*?</c:val>`)
	matches := re.FindAllSubmatch(fileContent, -1)

	if len(matches) == 0 {
		return fileContent, nil
	}

	for _, index := range indexes {
		fileContent = bytes.Replace(fileContent, []byte("<c:v>0</c:v>"), []byte(fmt.Sprintf("<c:v>%s</c:v>", values[index])), 1)
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
