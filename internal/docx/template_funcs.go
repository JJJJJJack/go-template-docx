package docx

import (
	"fmt"
	"strings"
)

type XmlImageData struct {
	DocPrId uint32
	Name    string
	RefID   string
}

const imageTemplateXml = `<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <wp:extent cx="2489026" cy="2489026" />
    <wp:docPr id="{{.DocPrId}}" name="{{.Name}}" />
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="0" name="{{.Name}}" />
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

const (
	DOCX_NEWLINE_INJECT        = "</w:t><w:br/><w:t>"
	DOCX_BREAKPARAGRAPH_INJECT = "</w:t></w:r></w:p><w:p><w:r><w:t>"
	RGB_SHADING_WRAPPER_F      = `<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="%s"/></w:rPr><w:t>%s</w:t>`
)

func image(filename string) string {
	return fmt.Sprintf("[[IMAGE:%s]]", filename)
}

func replaceImage(filename string) string {
	return fmt.Sprintf("[[REPLACE_IMAGE:%s]]", filename)
}

// preserveNewline newlines are treated as `SHIFT + ENTER` input,
// thus keeping the text in the same paragraph.
func preserveNewline(s string) string {
	return strings.ReplaceAll(s, "\n", DOCX_NEWLINE_INJECT)
}

// breakParagraph newlines are treated as `ENTER` input,
// thus creating a new paragraph for the sequent line.
func breakParagraph(s string) string {
	return strings.ReplaceAll(s, "\n", DOCX_BREAKPARAGRAPH_INJECT)
}

// shadeTextBg applies a background color to the given text
func shadeTextBg(hex, s string) string {
	return fmt.Sprintf(RGB_SHADING_WRAPPER_F, hex, s)
}

// shapeBgFillColor replace fillcolor to shapes
func shapeBgFillColor(s string) string {
	return fmt.Sprintf("[[SHAPE_BG_FILL_COLOR:%s]]", s)
}

// tableCellBgColor replace background color of table cells
func tableCellBgColor(s string) string {
	return fmt.Sprintf("[[TABLE_CELL_BG_COLOR:%s]]", s)
}
