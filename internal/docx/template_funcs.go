package docx

import (
	"fmt"
	"strconv"
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
	DOCX_NEWLINE_INJECT        = `</w:t><w:br/><w:t>`
	DOCX_BREAKPARAGRAPH_INJECT = `</w:t></w:r></w:p><w:p><w:r><w:t>`
	RGB_SHADING_WRAPPER_F      = `<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="%s"/></w:rPr><w:t>%s</w:t>`

	STYLE_WRAPPER_F     = `<w:rPr>%s</w:rPr><w:t>%s</w:t>`
	BOLD_W_TAG          = `<w:b /><w:bCs />`
	ITALIC_W_TAG        = `<w:i /><w:iCs />`
	UNDERLINE_W_TAG     = `<w:u w:val="single"/>`
	STRIKETHROUGH_W_TAG = `<w:strike />`
	FONT_SIZE_W_TAGS    = `<w:sz w:val="%d" /><w:szCs w:val="%d" />`
)

var (
	BOLD_WRAPPER_F          = fmt.Sprintf(STYLE_WRAPPER_F, BOLD_W_TAG, "%s")
	ITALIC_WRAPPER_F        = fmt.Sprintf(STYLE_WRAPPER_F, ITALIC_W_TAG, "%s")
	UNDERLINE_WRAPPER_F     = fmt.Sprintf(STYLE_WRAPPER_F, UNDERLINE_W_TAG, "%s")
	STRIKETHROUGH_WRAPPER_F = fmt.Sprintf(STYLE_WRAPPER_F, STRIKETHROUGH_W_TAG, "%s")
)

func fontSizeWrapperf(sizeHalfPoints int) string {
	if sizeHalfPoints <= 0 {
		sizeHalfPoints = 1
	}

	return fmt.Sprintf(FONT_SIZE_W_TAGS, sizeHalfPoints*2, sizeHalfPoints*2)
}

const (
	FONT_SIZE_STYLE_PREFIX       = "fontsize:"
	FONT_SIZE_STYLE_PREFIX_SHORT = "fs:"
)

// styledText applies multiple styles to the given text.
// The first argument is the text, the following arguments are styles.
func styledText(s ...string) (string, error) {
	if len(s) < 2 {
		return "", fmt.Errorf("styledText requires at least 1 text argument followed by style arguments")
	}

	text := ""
	styles := ""
	for i := range s {
		if i == 0 {
			text = s[i]
			continue
		}

		styleParam := strings.ToLower(s[i])

		if strings.HasPrefix(styleParam, FONT_SIZE_STYLE_PREFIX) || strings.HasPrefix(styleParam, FONT_SIZE_STYLE_PREFIX_SHORT) {
			if strings.Contains(styles, "<w:sz w:val=") {
				return "", fmt.Errorf("styledText got multiple font size styles")
			}

			sizeStr := strings.TrimPrefix(styleParam, FONT_SIZE_STYLE_PREFIX)
			sizeStr = strings.TrimPrefix(sizeStr, FONT_SIZE_STYLE_PREFIX_SHORT)

			ptSize, err := strconv.Atoi(sizeStr)
			if err != nil {
				return "", fmt.Errorf("styledText got invalid size: %s", sizeStr)
			}

			styles += fontSizeWrapperf(ptSize)
			continue
		}

		switch styleParam {
		case "b", "bold":
			if strings.Contains(styles, BOLD_W_TAG) {
				return "", fmt.Errorf("styledText got multiple bold styles")
			}

			styles += BOLD_W_TAG
		case "i", "italic":
			if strings.Contains(styles, ITALIC_W_TAG) {
				return "", fmt.Errorf("styledText got multiple italic styles")
			}

			styles += ITALIC_W_TAG
		case "u", "underline":
			if strings.Contains(styles, UNDERLINE_W_TAG) {
				return "", fmt.Errorf("styledText got multiple underline styles")
			}

			styles += UNDERLINE_W_TAG
		case "s", "strike", "strikethrough":
			if strings.Contains(styles, STRIKETHROUGH_W_TAG) {
				return "", fmt.Errorf("styledText got multiple strikethrough styles")
			}

			styles += STRIKETHROUGH_W_TAG
		default:
			return "", fmt.Errorf("styledText got unknown style: %s", s[i])
		}
	}

	return fmt.Sprintf(STYLE_WRAPPER_F, styles, text), nil
}

// bold makes the text bold
func bold(s string) string {
	return fmt.Sprintf(BOLD_WRAPPER_F, s)
}

// italic makes the text italic
func italic(s string) string {
	return fmt.Sprintf(ITALIC_WRAPPER_F, s)
}

// underline underlines the text
func underline(s string) string {
	return fmt.Sprintf(UNDERLINE_WRAPPER_F, s)
}

// strike applies strikethrough to the text
func strike(s string) string {
	return fmt.Sprintf(STRIKETHROUGH_WRAPPER_F, s)
}

// fontSize sets the font size of the text
func fontSize(s string, sizeHalfPoints int) string {
	return fmt.Sprintf(STYLE_WRAPPER_F, fontSizeWrapperf(sizeHalfPoints), s)
}

// image wraps a placeholder around the given filename for image insertion in the document.
func image(filename string) string {
	return fmt.Sprintf("[[IMAGE:%s]]", filename)
}

// replaceImage insert a placeholder around the given filename for image replacement in the document.
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
