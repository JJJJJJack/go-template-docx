package docx

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
	"text/template"
)

type XmlImageData struct {
	DocPrId uint32
	Name    string
	RefID   string
	Cx      int
	Cy      int
	// BorderColor is the RRGGBB border color, empty means no border
	BorderColor    string
	BorderWidthEmu int
	// BorderExtentEmu is the space reserved around the image for the border,
	// Word draws picture borders entirely outside the image so it must equal the border width
	// otherwise the border gets clipped by the previous line or covered by adjacent images
	BorderExtentEmu int
}

const imageTemplateXml = `<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <wp:extent cx="{{.Cx}}" cy="{{.Cy}}" />
    {{- if .BorderColor}}
    <wp:effectExtent l="{{.BorderExtentEmu}}" t="{{.BorderExtentEmu}}" r="{{.BorderExtentEmu}}" b="{{.BorderExtentEmu}}" />
    {{- end}}
    <wp:docPr id="{{.DocPrId}}" name="{{.Name}}" />
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{{.DocPrId}}" name="{{.Name}}" />
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
              <a:ext cx="{{.Cx}}" cy="{{.Cy}}" />
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst />
            </a:prstGeom>
            {{- if .BorderColor}}
            <a:ln w="{{.BorderWidthEmu}}">
              <a:solidFill>
                <a:srgbClr val="{{.BorderColor}}" />
              </a:solidFill>
              <a:miter lim="800000" />
            </a:ln>
            {{- end}}
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`

const (
	DOCX_NEWLINE_INJECT        = `</w:t><w:br/><w:t>`
	DOCX_BREAKPARAGRAPH_INJECT = `</w:t></w:r></w:p><w:p><w:r><w:t>`

	STYLE_WRAPPER_F     = `<w:rPr>%s</w:rPr><w:t>%s</w:t>`
	BOLD_W_TAG          = `<w:b /><w:bCs />`
	ITALIC_W_TAG        = `<w:i /><w:iCs />`
	UNDERLINE_W_TAG     = `<w:u w:val="single"/>`
	STRIKETHROUGH_W_TAG = `<w:strike />`
	FONT_SIZE_W_TAGS_F  = `<w:sz w:val="%d" /><w:szCs w:val="%d" />`
	COLOR_W_TAG_F       = `<w:color w:val="%s" />`
	HIGHLIGHT_W_TAG_F   = `<w:highlight w:val="%s" />`
	// HIGHLIGHT all values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.highlightcolor?view=openxml-2.8.1
	SHADING_W_TAG_F = `<w:shd w:val="clear" w:color="auto" w:fill="%s"/>`
)

var (
	BOLD_WRAPPER_F          = fmt.Sprintf(STYLE_WRAPPER_F, BOLD_W_TAG, "%s")
	ITALIC_WRAPPER_F        = fmt.Sprintf(STYLE_WRAPPER_F, ITALIC_W_TAG, "%s")
	UNDERLINE_WRAPPER_F     = fmt.Sprintf(STYLE_WRAPPER_F, UNDERLINE_W_TAG, "%s")
	STRIKETHROUGH_WRAPPER_F = fmt.Sprintf(STYLE_WRAPPER_F, STRIKETHROUGH_W_TAG, "%s")
	COLOR_WRAPPER_F         = fmt.Sprintf(STYLE_WRAPPER_F, COLOR_W_TAG_F, "%s")
	HIGHLIGHT_WRAPPER_F     = fmt.Sprintf(STYLE_WRAPPER_F, HIGHLIGHT_W_TAG_F, "%s")
	SHADING_WRAPPER_F       = fmt.Sprintf(STYLE_WRAPPER_F, SHADING_W_TAG_F, "%s")
)

func fontSizeWrapperf(sizeHalfPoints int) string {
	if sizeHalfPoints <= 0 {
		sizeHalfPoints = 1
	}

	return fmt.Sprintf(FONT_SIZE_W_TAGS_F, sizeHalfPoints*2, sizeHalfPoints*2)
}

const (
	FONT_SIZE_STYLE_PREFIX       = "fontSize:"
	FONT_SIZE_STYLE_PREFIX_SHORT = "fs:"
	TEXT_SHADING_STYLE_PREFIX    = "bg:"
)

// list enables you to take a variadic number of arguments and
// returns them as a slice of interface{} to another function
// directly from the template expressions.
func list(args ...interface{}) []interface{} {
	return args
}

// formatStylesTags takes a slice of styles and returns the corresponding XML tags.
func formatStylesTags(stylesList []interface{}, funcName string) (string, error) {
	styles := ""
	for _, arg := range stylesList {
		styleParam, ok := arg.(string)
		if !ok {
			return "", fmt.Errorf("%s got non-string style parameter: %v", funcName, arg)
		}

		// font size style
		if strings.HasPrefix(styleParam, FONT_SIZE_STYLE_PREFIX) || strings.HasPrefix(styleParam, FONT_SIZE_STYLE_PREFIX_SHORT) {
			if strings.Contains(styles, "<w:sz w:val=") {
				return "", fmt.Errorf("%s got multiple font size styles", funcName)
			}

			sizeStr := strings.TrimPrefix(styleParam, FONT_SIZE_STYLE_PREFIX)
			sizeStr = strings.TrimPrefix(sizeStr, FONT_SIZE_STYLE_PREFIX_SHORT)

			ptSize, err := strconv.Atoi(sizeStr)
			if err != nil {
				return "", fmt.Errorf("%s got invalid size: %s", funcName, sizeStr)
			}

			styles += fontSizeWrapperf(ptSize)
			continue
		}

		// color style
		if strings.HasPrefix(styleParam, "#") {
			if strings.Contains(styles, "<w:color w:val=") {
				return "", fmt.Errorf("%s got multiple color styles", funcName)
			}

			hex := strings.ToUpper(strings.TrimPrefix(styleParam, "#"))

			styles += fmt.Sprintf(COLOR_W_TAG_F, hex)
			continue
		}

		// shading style
		if strings.HasPrefix(styleParam, TEXT_SHADING_STYLE_PREFIX) {
			if strings.Contains(styles, "<w:shd w:val=") {
				return "", fmt.Errorf("%s got multiple background shading styles", funcName)
			}

			hex := strings.ToUpper(strings.TrimPrefix(styleParam, TEXT_SHADING_STYLE_PREFIX))
			hex = strings.TrimPrefix(hex, "#")

			styles += fmt.Sprintf(SHADING_W_TAG_F, hex)
			continue
		}

		switch styleParam {
		case "b", "bold":
			if strings.Contains(styles, BOLD_W_TAG) {
				return "", fmt.Errorf("%s got multiple bold styles", funcName)
			}

			styles += BOLD_W_TAG
		case "i", "italic":
			if strings.Contains(styles, ITALIC_W_TAG) {
				return "", fmt.Errorf("%s got multiple italic styles", funcName)
			}

			styles += ITALIC_W_TAG
		case "u", "underline":
			if strings.Contains(styles, UNDERLINE_W_TAG) {
				return "", fmt.Errorf("%s got multiple underline styles", funcName)
			}

			styles += UNDERLINE_W_TAG
		case "s", "strike", "strikethrough":
			if strings.Contains(styles, STRIKETHROUGH_W_TAG) {
				return "", fmt.Errorf("%s got multiple strikethrough styles", funcName)
			}

			styles += STRIKETHROUGH_W_TAG
		case "black", "blue", "cyan", "green",
			"magenta", "red", "yellow", "white",
			"darkBlue", "darkCyan", "darkGreen",
			"darkMagenta", "darkRed", "darkYellow",
			"darkGray", "lightGray", "none":
			if strings.Contains(styles, "<w:highlight w:val=") {
				return "", fmt.Errorf("%s got multiple highlight colors styles", funcName)
			}

			styles += fmt.Sprintf(HIGHLIGHT_W_TAG_F, styleParam)
		default:
			return "", fmt.Errorf("%s got unknown style: %s", funcName, styleParam)
		}
	}

	return styles, nil
}

// styledText takes a strings and a slice of styles to apply to the text.
// You can use this function to style text with a set variable containing
// a reusable style in your code.
func styledText(text string, styles []interface{}) (string, error) {
	stylesTags, err := formatStylesTags(styles, "styledText")
	if err != nil {
		return "", err
	}

	return fmt.Sprintf(STYLE_WRAPPER_F, stylesTags, text), nil
}

// inlineStyledText applies multiple styles to the given text.
// The first argument is the text, the following arguments are styles.
// You can use this function to apply multiple styles to a text without
// having to wrap them in a list.
func inlineStyledText(text string, styles ...interface{}) (string, error) {
	stylesTags, err := formatStylesTags(styles, "inlineStyledText")
	if err != nil {
		return "", err
	}

	return fmt.Sprintf(STYLE_WRAPPER_F, stylesTags, text), nil
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
func fontSize(s string, size int) string {
	return fmt.Sprintf(STYLE_WRAPPER_F, fontSizeWrapperf(size), s)
}

// color sets the font color of the text
func color(s, hex string) (string, error) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "", fmt.Errorf("func 'color': invalid hex color value: %s (must be 6 characters like '0077FF')", hex)
	}

	return fmt.Sprintf(COLOR_WRAPPER_F, strings.ToUpper(hex), s), nil
}

// highlight applies a highlight color to the text
func highlight(s, color string) (string, error) {
	switch color {
	case "black":
	case "blue":
	case "cyan":
	case "green":
	case "magenta":
	case "red":
	case "yellow":
	case "white":
	case "darkBlue":
	case "darkCyan":
	case "darkGreen":
	case "darkMagenta":
	case "darkRed":
	case "darkYellow":
	case "darkGray":
	case "lightGray":
	case "none":
	default:
		return "", fmt.Errorf("func 'highlight': invalid highlight color value: %s", color)
	}

	return fmt.Sprintf(HIGHLIGHT_WRAPPER_F, color, s), nil
}

// shadeTextBg applies a background color to the given text
func shadeTextBg(s, hex string) (string, error) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "", fmt.Errorf("func 'shadeTextBg': invalid hex color value: %s (must be 6 characters like '0077FF')", hex)
	}

	return fmt.Sprintf(SHADING_WRAPPER_F, strings.ToUpper(hex), s), nil
}

const (
	IMAGE_BORDER_COLOR_PREFIX       = "borderColor:"
	IMAGE_BORDER_COLOR_PREFIX_SHORT = "bc:"
	IMAGE_BORDER_WIDTH_PREFIX       = "borderWidth:"
	IMAGE_BORDER_WIDTH_PREFIX_SHORT = "bw:"

	defaultImageBorderWidthPt = 1.0
	// maximum line width allowed by the DrawingML schema (20116800 EMUs)
	maxImageBorderWidthPt = 1584.0
	emusPerPoint          = 12700
)

var hexColorRe = regexp.MustCompile(`^[0-9A-Fa-f]{6}$`)

type imageOptions struct {
	// BorderColor is the RRGGBB border color, empty means no border
	BorderColor   string
	BorderWidthPt float64
}

// validate checks the options consistency and applies the defaults.
func (o *imageOptions) validate(funcName string, borderWidthSet bool) error {
	if o.BorderColor != "" && !hexColorRe.MatchString(o.BorderColor) {
		return fmt.Errorf("%s got invalid border color: %s (must be 6 hex characters like '0077FF')", funcName, o.BorderColor)
	}
	o.BorderColor = strings.ToUpper(o.BorderColor)

	if !borderWidthSet {
		o.BorderWidthPt = defaultImageBorderWidthPt
	}

	if math.IsNaN(o.BorderWidthPt) || o.BorderWidthPt <= 0 || o.BorderWidthPt > maxImageBorderWidthPt {
		return fmt.Errorf("%s got invalid border width: %v (must be greater than 0 and at most %v points)", funcName, o.BorderWidthPt, maxImageBorderWidthPt)
	}

	return nil
}

// parseImageOptions parses the options passed to the image template function,
// nil and empty options are ignored so that they can be optionally set from the template data.
func parseImageOptions(opts []interface{}) (imageOptions, error) {
	options := imageOptions{}
	borderWidthSet := false

	for _, arg := range opts {
		if arg == nil {
			continue
		}

		opt, ok := arg.(string)
		if !ok {
			return options, fmt.Errorf("image got non-string option: %v", arg)
		}

		if opt == "" {
			continue
		}

		switch {
		case strings.HasPrefix(opt, IMAGE_BORDER_COLOR_PREFIX), strings.HasPrefix(opt, IMAGE_BORDER_COLOR_PREFIX_SHORT):
			options.BorderColor = strings.TrimPrefix(opt, IMAGE_BORDER_COLOR_PREFIX_SHORT)
			options.BorderColor = strings.TrimPrefix(options.BorderColor, IMAGE_BORDER_COLOR_PREFIX)
			options.BorderColor = strings.TrimPrefix(options.BorderColor, "#")
			if options.BorderColor == "" {
				return options, fmt.Errorf("image got empty border color")
			}
		case strings.HasPrefix(opt, IMAGE_BORDER_WIDTH_PREFIX), strings.HasPrefix(opt, IMAGE_BORDER_WIDTH_PREFIX_SHORT):
			if borderWidthSet {
				return options, fmt.Errorf("image got multiple border widths")
			}

			widthStr := strings.TrimPrefix(opt, IMAGE_BORDER_WIDTH_PREFIX)
			widthStr = strings.TrimPrefix(widthStr, IMAGE_BORDER_WIDTH_PREFIX_SHORT)

			width, err := strconv.ParseFloat(widthStr, 64)
			if err != nil {
				return options, fmt.Errorf("image got invalid border width: %s", widthStr)
			}

			options.BorderWidthPt = width
			borderWidthSet = true
			if options.BorderColor == "" {
				options.BorderColor = "000000"
			}
		default:
			return options, fmt.Errorf("image got unknown option: %s", opt)
		}
	}

	return options, options.validate("image", borderWidthSet)
}

// encode serializes the options inside the image placeholder,
// e.g. [[IMAGE[border=FF0000;bw=1.5]:filename.png]], the options are placed
// before the filename so that any character can be used in the filename.
func (o imageOptions) encode() string {
	if o.BorderColor == "" {
		return ""
	}

	return fmt.Sprintf("[border=%s;bw=%s]", o.BorderColor, strconv.FormatFloat(o.BorderWidthPt, 'f', -1, 64))
}

// decodeImageOptions parses the options serialized by imageOptions.encode,
// without the surrounding brackets.
func decodeImageOptions(s string) (imageOptions, error) {
	options := imageOptions{}
	borderWidthSet := false

	if s == "" {
		return options, nil
	}

	for _, kv := range strings.Split(s, ";") {
		key, value, found := strings.Cut(kv, "=")
		if !found {
			return options, fmt.Errorf("invalid image placeholder option: %s", kv)
		}

		switch key {
		case "border":
			options.BorderColor = value
		case "bw":
			width, err := strconv.ParseFloat(value, 64)
			if err != nil {
				return options, fmt.Errorf("invalid image placeholder border width: %s", value)
			}

			options.BorderWidthPt = width
			borderWidthSet = true
		default:
			return options, fmt.Errorf("unknown image placeholder option: %s", key)
		}
	}

	return options, options.validate("image placeholder", borderWidthSet)
}

// image wraps a placeholder around the given filename for image insertion in the document.
// Optional styles can be passed after the filename:
//   - "border:#RRGGBB" adds a border with the given color
//   - "borderWidth:N" or "bw:N" sets the border width in points (default 1)
//
// nil and empty options are ignored.
func image(filename string, opts ...interface{}) (string, error) {
	options, err := parseImageOptions(opts)
	if err != nil {
		return "", err
	}

	return fmt.Sprintf("[[IMAGE%s:%s]]", options.encode(), filename), nil
}

// replaceImage insert a placeholder around the given filename for image replacement in the document.
func replaceImage(filename string) string {
	return fmt.Sprintf("[[REPLACE_IMAGE:%s]]", filename)
}

// preserveNewline newlines are treated as `SHIFT + ENTER` input,
// thus keeping the text in the same paragraph.
func preserveNewline(text string) string {
	return strings.ReplaceAll(text, "\n", DOCX_NEWLINE_INJECT)
}

// breakParagraph newlines are treated as `ENTER` input,
// thus creating a new paragraph for the sequent line.
func breakParagraph(text string) string {
	return strings.ReplaceAll(text, "\n", DOCX_BREAKPARAGRAPH_INJECT)
}

// shapeBgFillColor replace fillcolor to shapes
func shapeBgFillColor(hex string) (string, error) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "", fmt.Errorf("func 'shapeBgFillColor': invalid hex color value: %s  (must be 6 characters like '0077FF')", hex)
	}

	return fmt.Sprintf("[[SHAPE_BG_FILL_COLOR:%s]]", strings.ToUpper(hex)), nil
}

// tableCellBgColor replace background color of table cells
func tableCellBgColor(hex string) (string, error) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return "", fmt.Errorf("func 'tableCellBgColor': invalid hex color value: %s  (must be 6 characters like '0077FF')", hex)
	}

	return fmt.Sprintf("[[TABLE_CELL_BG_COLOR:%s]]", strings.ToUpper(hex)), nil
}

var TemplateFuncs = template.FuncMap{
	"list":             list,
	"bold":             bold,
	"italic":           italic,
	"underline":        underline,
	"strike":           strike,
	"fontSize":         fontSize,
	"inlineStyledText": inlineStyledText,
	"styledText":       styledText,
	"color":            color,
	"highlight":        highlight,
	"preserveNewline":  preserveNewline,
	"breakParagraph":   breakParagraph,
	"shadeTextBg":      shadeTextBg,
	"image":            image,
	"replaceImage":     replaceImage,
	"shapeBgFillColor": shapeBgFillColor,
	"tableCellBgColor": tableCellBgColor,
}
