package docx

import (
	"bytes"
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
	"text/template"
)

func (d *documentMeta) applyImages(srcXML string) (string, []MediaRel, error) {
	mediaRels := []MediaRel{}

	imagePlaceholderRE := regexp.MustCompile(`\[\[IMAGE:.*?\]\]`)
	xmlBlocks := imagePlaceholderRE.FindAllString(srcXML, -1)
	for _, xmlBlock := range xmlBlocks {
		filename := strings.TrimPrefix(xmlBlock, "[[IMAGE:")
		filename = strings.TrimSuffix(filename, "]]")

		buffer := bytes.Buffer{}
		docPrId, err := d.RandUniqueDocPrId()
		if err != nil {
			return srcXML, mediaRels, fmt.Errorf("unable to get unique docPrId: %w", err)
		}

		rid := d.NextRId()
		rId := fmt.Sprintf("rId%d", rid)

		imageTemplate, err := template.New("image-template").Parse(imageTemplateXml)
		if err != nil {
			return srcXML, mediaRels, err
		}

		err = imageTemplate.Execute(&buffer, XmlImageData{
			DocPrId: docPrId,
			Name:    filename,
			RefID:   rId,
		})
		if err != nil {
			return srcXML, mediaRels, fmt.Errorf("unable to execute image template: %w", err)
		}

		mediaRels = append(mediaRels, MediaRel{
			Type:   ImageMediaType,
			RefID:  rId,
			Source: path.Join("media", filename),
		})

		srcXML = strings.ReplaceAll(srcXML, xmlBlock, buffer.String())
	}

	return srcXML, mediaRels, nil
}

// replaceImages looks for [[REPLACE_IMAGE:filename.ext]] placeholders inside <w:drawing>...</w:drawing> blocks
// remove the placeholder and replaces the image reference inside the block with the given image's rId.
func (d *documentMeta) replaceImages(srcXML string) (string, []MediaRel) {
	anchorRe := regexp.MustCompile(`(?s)<w:drawing>.*?</w:drawing>`)
	placeholderRe := regexp.MustCompile(`\[\[REPLACE_IMAGE:([^\]]+)\]\]`)
	blipRe := regexp.MustCompile(`(<a:blip\s+r:embed=")[^"]*(")`)

	mediaRels := []MediaRel{}

	result := anchorRe.ReplaceAllStringFunc(srcXML, func(block string) string {
		pm := placeholderRe.FindStringSubmatch(block)
		if len(pm) < 2 {
			return block
		}
		filename := pm[1]

		block = placeholderRe.ReplaceAllString(block, "")

		rid := d.NextRId()
		rId := fmt.Sprintf("rId%d", rid)

		mediaRels = append(mediaRels, MediaRel{
			Type:   ImageMediaType,
			RefID:  rId,
			Source: path.Join("media", filename),
		})

		block = blipRe.ReplaceAllString(block, "${1}"+rId+"${2}")

		return block
	})

	return result, mediaRels
}

const withoutSolidFill = `</a:prstGeom></wps:spPr>`
const withSolidFill = `</a:prstGeom><a:solidFill><a:srgbClr val="ffffff" /></a:solidFill></wps:spPr>`

// ReplaceAllShapeBgColors finds shapes that contain the [[SHAPE_BG_COLOR:RRGGBB]]/[[SHAPE_BG_COLOR:#RRGGBB]]
// placeholder and uses its value to replace the fillcolor attribute of the shape
// TODO: replace with proper XML parsing
func (d *documentMeta) applyShapesBgFillColor(srcXML string) string {
    placeholderRe := regexp.MustCompile(`\[\[SHAPE_BG_FILL_COLOR:#?([0-9A-Fa-f]{6})\]\]`)

    altContentRe := regexp.MustCompile(`(?s)<mc:AlternateContent>.*?</mc:AlternateContent>`)

    return altContentRe.ReplaceAllStringFunc(srcXML, func(block string) string {
        placeholders := placeholderRe.FindAllStringSubmatch(block, -1)
        if len(placeholders) == 0 {
            return block
        }

        for _, pm := range placeholders {
            hex := strings.ToLower(pm[1])

            // 1) OOXML path (wps:spPr)
            // If there is a gradient fill, keep it but recolor stops to srgbClr with target hex
            gradFillRe := regexp.MustCompile(`(?is)<a:gradFill\b[\s\S]*?</a:gradFill>`)
            block = gradFillRe.ReplaceAllStringFunc(block, func(grad string) string {
                // Convert schemeClr -> srgbClr, preserving nested transforms (shade, satMod, etc.)
                schemeToSrgb := regexp.MustCompile(`(?is)<a:schemeClr\b[^>]*>([\s\S]*?)</a:schemeClr>`)
                grad = schemeToSrgb.ReplaceAllString(grad, `<a:srgbClr val="`+hex+`">$1</a:srgbClr>`)

                // Ensure any existing srgbClr val is set to hex
                srgbValAttrRe := regexp.MustCompile(`(?is)(<a:srgbClr\b[^>]*\bval=")[^"]*(")`)
                grad = srgbValAttrRe.ReplaceAllString(grad, `${1}`+hex+`${2}`)
                return grad
            })

            // If no gradFill present, fall back to solid fill handling
            if !strings.Contains(block, "<a:gradFill") {
                // Update existing solid fill color if present
                srgbValAttrRe := regexp.MustCompile(`(?is)(<a:srgbClr\b[^>]*\bval=")[^"]*(")`)
                if srgbValAttrRe.MatchString(block) {
                    block = srgbValAttrRe.ReplaceAllString(block, `${1}`+hex+`${2}`)
                } else {
                    // Insert a solid fill after </a:prstGeom> if none exists
                    prstGeomCloseRe := regexp.MustCompile(`(?is)</a:prstGeom>\s*`)
                    if loc := prstGeomCloseRe.FindStringIndex(block); loc != nil {
                        // Insert just after </a:prstGeom>
                        insert := `<a:solidFill><a:srgbClr val="` + hex + `" /></a:solidFill>`
                        block = block[:loc[1]] + insert + block[loc[1]:]
                    } else {
                        // Fallback to legacy pattern replacement
                        block = strings.Replace(block, withoutSolidFill, `</a:prstGeom><a:solidFill><a:srgbClr val="`+hex+`" /></a:solidFill></wps:spPr>`, 1)
                    }
                }
            }

            // 2) VML fallback (<v:shape>): update fillcolor and convert gradient <v:fill/> to solid
            fillColorRe := regexp.MustCompile(`(?i)\bfillcolor="[^"]*"`)
            block = fillColorRe.ReplaceAllStringFunc(block, func(fc string) string {
                return `fillcolor="#` + hex + `"`
            })

            vmlFillRe := regexp.MustCompile(`(?is)<v:fill\b[^>]*/>`) // update VML gradient fill, keep gradient and recolor stops
            block = vmlFillRe.ReplaceAllStringFunc(block, func(tag string) string {
                // Extract commonly used attrs to preserve
                get := func(name string) string {
                    re := regexp.MustCompile(`(?i)\b` + name + `="([^"]*)"`)
                    m := re.FindStringSubmatch(tag)
                    if len(m) == 2 { return m[1] }
                    return ""
                }
                rotate := get("rotate")
                angle := get("angle")
                focus := get("focus")

                // Compute simple darker/lighter variants to keep a visible gradient
                base := strings.ToLower(strings.TrimPrefix(hex, "#"))
                darker := adjustBrightnessHex(base, 0.25, false)
                lighter := adjustBrightnessHex(base, 0.35, true)

                attrs := []string{`type="gradient"`, `colors="0 #`+darker+`;0.5 #`+base+`;1 #`+lighter+`"`, `color2="#`+lighter+`"`}
                if rotate != "" { attrs = append(attrs, `rotate="`+rotate+`"`) }
                if angle != "" { attrs = append(attrs, `angle="`+angle+`"`) }
                if focus != "" { attrs = append(attrs, `focus="`+focus+`"`) }
                return `<v:fill ` + strings.Join(attrs, " ") + `/>`
            })
        }

        block = placeholderRe.ReplaceAllString(block, "")

        return block
    })
}

// adjustBrightnessHex lightens or darkens a hex color by factor (0..1).
// If lighten is true, moves towards 255; else towards 0.
func adjustBrightnessHex(hex string, factor float64, lighten bool) string {
    if len(hex) != 6 { return hex }
    r, _ := strconv.ParseUint(hex[0:2], 16, 8)
    g, _ := strconv.ParseUint(hex[2:4], 16, 8)
    b, _ := strconv.ParseUint(hex[4:6], 16, 8)
    rf, gf, bf := float64(r), float64(g), float64(b)
    if lighten {
        rf = rf + (255.0-rf)*factor
        gf = gf + (255.0-gf)*factor
        bf = bf + (255.0-bf)*factor
    } else {
        rf = rf * (1.0 - factor)
        gf = gf * (1.0 - factor)
        bf = bf * (1.0 - factor)
    }
    clamp := func(x float64) uint8 {
        if x < 0 { x = 0 }
        if x > 255 { x = 255 }
        return uint8(x + 0.5)
    }
    return fmt.Sprintf("%02x%02x%02x", clamp(rf), clamp(gf), clamp(bf))
}

const withoutShading = `></w:tcPr>`
const withShading = `><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" /></w:tcPr>`

// replaceTableCellBgColors is used to apply the hex color found in the
// [[TABLE_CELL_BG_COLOR:RRGGBB]]/[[TABLE_CELL_BG_COLOR:#RRGGBB]] as the background color of the table cell
func (d *documentMeta) replaceTableCellBgColors(srcXML string) string {
	tcRe := regexp.MustCompile(`(?s)<w:tc>.*?</w:tc>`)

	output := tcRe.ReplaceAllStringFunc(srcXML, func(block string) string {
		hexRe := regexp.MustCompile(`\[\[TABLE_CELL_BG_COLOR:#?([0-9A-Fa-f]{6})\]\]`)
		hexMatch := hexRe.FindStringSubmatch(block)
		if len(hexMatch) < 2 {
			return block
		}
		hex := hexMatch[1]

		if !regexp.MustCompile(`(?i)<w:shd[^>]*?/>`).MatchString(block) {
			block = strings.Replace(block, withoutShading, withShading, 1)
		}

		fillRe := regexp.MustCompile(`(?i)(<w:shd[^>]*? w:fill=")[^"]*(")`)
		if fillRe.MatchString(block) {
			block = fillRe.ReplaceAllString(block, `${1}`+hex+`${2}`)
		} else {
			shdRe := regexp.MustCompile(`(?i)(<w:shd)`)
			block = shdRe.ReplaceAllString(block, `${1} w:fill="`+hex+`"`)
		}

		block = hexRe.ReplaceAllString(block, ``)
		block = strings.ReplaceAll(block, "<w:r><w:t></w:t></w:r>", "")

		return block
	})

	return output
}
