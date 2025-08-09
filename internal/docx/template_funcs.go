package docx

import (
	"bytes"
	"fmt"
	"path"
	"regexp"
	"strings"
	"text/template"
)

var counterID = 100

type ImageData struct {
	ID    int
	Name  string
	RefID string
}

const imageTemplate = `<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <wp:extent cx="2489026" cy="2489026" />
    <wp:docPr id="{{.ID}}" name="{{.Name}}" />
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{{.ID}}" name="{{.Name}}" />
            <pic:cNvPicPr />
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{{.RefID}}" />
            <a:stretch>
              <a:fillRect />
            </a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0" />
              <a:ext cx="2489026" cy="2489026" />
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst />
            </a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`

func toImage(s string) string {
	return fmt.Sprintf("[[IMAGE:%s]]", s)
}

func generateSequentialReferenceID() string {
	counterID++

	return fmt.Sprintf("rId%d", counterID)
}

func applyImages(srcXML string) (string, []MediaRel, error) {
	mediaList := []MediaRel{}

	re := regexp.MustCompile(`<w:[A-Za-z]?>\[\[IMAGE:.*?\]\]</w:[A-Za-z]>`)
	imagePlaceholderRE := regexp.MustCompile(`\[\[IMAGE:.*?\]\]`)
	xmlBlocks := re.FindAllString(srcXML, -1)
	for _, xmlBlock := range xmlBlocks {
		imageTemplate, err := template.New("image-template").Parse(imageTemplate)
		if err != nil {
			return srcXML, mediaList, err
		}

		imageDirections := imagePlaceholderRE.FindAllString(xmlBlock, -1)
		if len(imageDirections) < 1 {
			continue
		}

		filename := strings.TrimPrefix(imageDirections[0], "[[IMAGE:")
		filename = strings.TrimSuffix(filename, "]]")

		buffer := bytes.Buffer{}
		refID := generateSequentialReferenceID()

		imageTemplate.Execute(&buffer, ImageData{
			ID:    counterID,
			Name:  filename,
			RefID: refID,
		})

		mediaList = append(mediaList, MediaRel{
			Type:   ImageMediaType,
			RefID:  refID,
			Source: path.Join("media", filename),
		})

		srcXML = strings.ReplaceAll(srcXML, xmlBlock, buffer.String())
	}

	return srcXML, mediaList, nil
}
