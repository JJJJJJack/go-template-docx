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

const imageCenterTemplateXml = `<w:p w14:paraId="4E9843FD" w14:textId="77777777" w:rsidR="00963E01" w:rsidRDefault="0028620A"
      w:rsidP="00963E01">
      <w:pPr>
        <w:keepNext />
        <w:spacing w:after="0" w:line="240" w:lineRule="auto" />
        <w:jc w:val="center" />
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsia="Calibri" />
          <w:noProof />
          <w:u w:val="single" />
        </w:rPr>
        <w:lastRenderedPageBreak />
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="68DC1111"
            wp14:editId="381D76A9">
            <wp:extent cx="2543175" cy="2962275" />
            <wp:effectExtent l="0" t="0" r="0" b="0" />
            <wp:docPr id="1228041855" name="Picture 1" />
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                noChangeAspect="1" />
            </wp:cNvGraphicFramePr>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Picture 1" />
                    <pic:cNvPicPr>
                      <a:picLocks noChangeAspect="1" noChangeArrowheads="1" />
                    </pic:cNvPicPr>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId10">
                      <a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                          <a14:useLocalDpi
                            xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                            val="0" />
                        </a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:srcRect />
                    <a:stretch>
                      <a:fillRect />
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr bwMode="auto">
                    <a:xfrm>
                      <a:off x="0" y="0" />
                      <a:ext cx="2543175" cy="2962275" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst />
                    </a:prstGeom>
                    <a:noFill />
                    <a:ln>
                      <a:noFill />
                    </a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
		`

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
	DOCX_NEWLINE_INJECT        = "</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>"
	DOCX_BREAKPARAGRAPH_INJECT = "</w:t></w:r></w:p><w:p><w:r><w:t>"
	RGB_SHADING_WRAPPER_F      = `<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="%s"/></w:rPr><w:t>%s</w:t></w:r>`
)

func image(filename string) string {
	return fmt.Sprintf("[[IMAGE:%s]]", filename)
}

func toCenteredImage(s string) string {
	return fmt.Sprintf("[[CENTERED_IMAGE:%s]]", s)
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
	return fmt.Sprintf("[[SHAPE_BG_COLOR:%s]]", s)
}
