package docx

import (
	"bytes"
	"encoding/xml"
)

type tagDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type tagOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type contentTypes struct {
	XMLName   xml.Name      `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []tagDefault  `xml:"Default"`
	Overrides []tagOverride `xml:"Override"`
}

func ParseContentTypes(data []byte) (*contentTypes, error) {
	var ct contentTypes
	err := xml.Unmarshal(data, &ct)
	if err != nil {
		return nil, err
	}
	return &ct, nil
}

// AddDefaultUnique adds a default content type if it does not already exist in the list.
func (ct *contentTypes) AddDefaultUnique(extension, contentType string) {
	for _, d := range ct.Defaults {
		if d.Extension == extension && d.ContentType == contentType {
			return
		}
	}

	ct.Defaults = append(ct.Defaults, tagDefault{
		Extension:   extension,
		ContentType: contentType,
	})
}

// replaceEmptyTags replaces specific XML empty tags patterns.
func replaceEmptyTags(data []byte) []byte {
	data = bytes.ReplaceAll(data, []byte("></Default>"), []byte(" />"))
	data = bytes.ReplaceAll(data, []byte("></Override>"), []byte(" />"))
	return data
}

func (ct *contentTypes) ToXml() ([]byte, error) {
	output, err := xml.MarshalIndent(ct, "", "  ")
	if err != nil {
		return []byte{}, err
	}

	output = replaceEmptyTags(output)

	header := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	xmlBytes := make([]byte, 0, len(header)+len(output))

	xmlBytes = append(xmlBytes, header...)
	xmlBytes = append(xmlBytes, output...)

	return xmlBytes, nil
}
