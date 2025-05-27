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

type relationship struct {
	XMLName       xml.Name             `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []relationshipDetail `xml:"Relationship"`
}

func (r *relationship) addMediaToRels(media []mediaRel) {
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

func (r *relationship) addRelationship(relType, target, id string) {
	newRel := relationshipDetail{
		Type:   relType,
		Target: target,
		Id:     id,
	}
	r.Relationships = append(r.Relationships, newRel)
}

func (r *relationship) toXML() (string, error) {
	output, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return "", err
	}
	return xml.Header + string(output), nil
}

func parseRelationship(data []byte) (*relationship, error) {
	var relationships relationship
	err := xml.Unmarshal(data, &relationships)
	if err != nil {
		return nil, err
	}

	return &relationships, nil
}
