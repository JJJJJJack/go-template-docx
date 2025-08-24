package docx

import (
	"encoding/xml"
)

const (
	imageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

type relationshipDetail struct {
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
	Id     string `xml:"Id,attr"`
}

type Relationship struct {
	XMLName       xml.Name             `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []relationshipDetail `xml:"Relationship"`
}

func (r *Relationship) AddMediaToRels(media []MediaRel) {
	for _, m := range media {
		switch m.Type {
		case ImageMediaType:
			r.addRelationship(
				imageRelationship,
				m.Source,
				m.RefID,
			)
		}
	}
}

func (r *Relationship) addRelationship(relType, target, id string) {
	newRel := relationshipDetail{
		Type:   relType,
		Target: target,
		Id:     id,
	}

	r.Relationships = append(r.Relationships, newRel)
}

func (r *Relationship) ToXml() ([]byte, error) {
	output, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return []byte{}, err
	}

	header := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	xmlBytes := make([]byte, 0, len(header)+len(output))

	xmlBytes = append(xmlBytes, header...)
	xmlBytes = append(xmlBytes, output...)

	return xmlBytes, nil
}

func ParseRelationship(data []byte) (*Relationship, error) {
	var relationships Relationship
	err := xml.Unmarshal(data, &relationships)
	if err != nil {
		return nil, err
	}

	return &relationships, nil
}
