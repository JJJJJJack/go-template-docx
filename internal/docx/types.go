package docx

import (
	"encoding/xml"
)

type defaultContent struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type overrideContent struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type contentTypes struct {
	XMLName   xml.Name          `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []defaultContent  `xml:"Default"`
	Overrides []overrideContent `xml:"Override"`
}

func parseContentTypes(data []byte) (*contentTypes, error) {
	var ct contentTypes
	err := xml.Unmarshal(data, &ct)
	if err != nil {
		return nil, err
	}
	return &ct, nil
}

func (ct *contentTypes) ensureImageDefaults(ext string, mime string) {
	for _, d := range ct.Defaults {
		if d.Extension == ext {
			return
		}
	}
	ct.Defaults = append(ct.Defaults, defaultContent{
		Extension:   ext,
		ContentType: mime,
	})
}

func (ct *contentTypes) toXML() (string, error) {
	output, err := xml.MarshalIndent(ct, "", "  ")
	if err != nil {
		return "", err
	}
	return xml.Header + string(output), nil
}
