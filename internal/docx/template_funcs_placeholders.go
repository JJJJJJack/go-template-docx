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

		v, ok := d.mediaMap[filename]
		if !ok {
			return srcXML, mediaRels, fmt.Errorf("filename '%s' not found in loaded medias", filename)
		}

		cx, cy, err := d.computeDocxImageSize(v)
		if err != nil {
			return srcXML, mediaRels, fmt.Errorf("unable to compute image size for '%s': %w", filename, err)
		}

		err = imageTemplate.Execute(&buffer, XmlImageData{
			DocPrId: docPrId,
			Name:    filename,
			RefID:   rId,
			Cx:      cx,
			Cy:      cy,
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

// adjustBrightnessHex lightens or darkens a hex color by factor (0..1).
// If lighten is true, moves towards 255; else towards 0.
func adjustBrightnessHex(hex string, factor float64, lighten bool) string {
	if len(hex) != 6 {
		return hex
	}
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
		if x < 0 {
			x = 0
		}
		if x > 255 {
			x = 255
		}
		return uint8(x + 0.5)
	}
	return fmt.Sprintf("%02x%02x%02x", clamp(rf), clamp(gf), clamp(bf))
}

const withoutSolidFill = `</a:prstGeom></wps:spPr>`

var (
	wpsSpPrRe            = regexp.MustCompile(`(?is)<wps:spPr>.*?</wps:spPr>`)
	placeholderRe        = regexp.MustCompile(`\[\[SHAPE_BG_FILL_COLOR:#?([0-9A-Fa-f]{6})\]\]`)
	mcAlternateContentRe = regexp.MustCompile(`(?s)<mc:AlternateContent>.*?</mc:AlternateContent>`)
	transRe              = regexp.MustCompile(`(?is)<a:(?:lumMod|tint|satMod)[^>]*/>`)
	aGradFillRe          = regexp.MustCompile(`(?is)<a:gradFill\b[\s\S]*?</a:gradFill>`)
	aSchemeToSrgb        = regexp.MustCompile(`(?is)<a:schemeClr\b[^>]*>([\s\S]*?)</a:schemeClr>`)
	aSrgbValAttrRe       = regexp.MustCompile(`(?is)(<a:srgbClr\b[^>]*\bval=")[^"]*(")`)
	aSolidFillRe         = regexp.MustCompile(`(?is)<a:solidFill\b[\s\S]*?</a:solidFill>`)
	aPrstGeomCloseRe     = regexp.MustCompile(`(?is)</a:prstGeom>\s*`)
	aShadingRe           = regexp.MustCompile(`(?is)<a:(?:lumMod|lumOff|tint|shade|satMod)\b[^>]*/>`)
	wpsStyleFixRe        = regexp.MustCompile(`(?is)<wps:style>.*?</wps:style>`)
	fillColorRe          = regexp.MustCompile(`(?i)\bfillcolor="[^"]*"`)
	aSrgbClrRe           = regexp.MustCompile(`(?is)<a:srgbClr([^>]*)>\s*(</a:(?:fillRef|effectRef|fontRef)>)`)
	vFillRe              = regexp.MustCompile(`(?is)<v:fill\b[^>]*/>`)
)

// ReplaceAllShapeBgColors finds shapes that contain the [[SHAPE_BG_COLOR:RRGGBB]]/[[SHAPE_BG_COLOR:#RRGGBB]]
// placeholder and uses its value to replace the fillcolor attribute of the shape
// TODO: replace with proper XML parsing
func (d *documentMeta) applyShapesBgFillColor(srcXML string) string {
	return mcAlternateContentRe.ReplaceAllStringFunc(srcXML, func(block string) string {
		placeholders := placeholderRe.FindAllStringSubmatch(block, -1)
		if len(placeholders) == 0 {
			return block
		}

		for _, pm := range placeholders {
			hex := strings.ToLower(pm[1])

			// Work only inside <wps:spPr> to target fill and avoid touching line/effect refs
			block = wpsSpPrRe.ReplaceAllStringFunc(block, func(sppr string) string {
				// Capture existing transform children to preserve visual shading
				transforms := strings.Join(transRe.FindAllString(sppr, -1), "")
				if transforms == "" {
					// Fallback to common Word shape shading if none detected
					transforms = `<a:lumMod val="10000"/><a:tint val="66000"/><a:satMod val="160000"/>`
				}

				// Gradient fill: recolor stops to srgbClr=hex, keep transforms/direction
				sppr = aGradFillRe.ReplaceAllStringFunc(sppr, func(grad string) string {
					grad = aSchemeToSrgb.ReplaceAllString(grad, `<a:srgbClr val="`+hex+`">$1</a:srgbClr>`)
					grad = aSrgbValAttrRe.ReplaceAllString(grad, `${1}`+hex+`${2}`)
					return grad
				})

				// Solid fill: convert schemeClr to srgbClr and strip transforms for exact color
				if !strings.Contains(sppr, "<a:gradFill") {
					// Ignore existing schemeClr/srgbClr children and rewrite the whole solidFill block

					// Replace or insert <a:solidFill> with a clean srgb color (no shading transforms)
					replacement := `<a:solidFill><a:srgbClr val="` + hex + `"/></a:solidFill>`
					// Unconditionally replace any existing solidFill, then ensure one exists
					sppr = aSolidFillRe.ReplaceAllString(sppr, replacement)
					if !strings.Contains(sppr, `<a:solidFill>`) {
						if loc := aPrstGeomCloseRe.FindStringIndex(sppr); loc != nil {
							sppr = sppr[:loc[1]] + replacement + sppr[loc[1]:]
						} else {
							sppr = strings.Replace(sppr, withoutSolidFill, `</a:prstGeom>`+replacement+`</wps:spPr>`, 1)
						}
					}

					// Ensure no shading transforms remain inside the solid fill
					sppr = aShadingRe.ReplaceAllString(sppr, "")
				}

				// Do not force transforms back into the solid fill; keep clean solid color

				return sppr
			})

			// Fix style blocks that might have self-closing srgbClr incorrectly expanded
			block = wpsStyleFixRe.ReplaceAllStringFunc(block, func(style string) string {
				// Turn `<a:srgbClr ...></a:fillRef>` into `<a:srgbClr .../></a:fillRef>` etc.
				style = aSrgbClrRe.ReplaceAllString(style, `<a:srgbClr$1/>$2`)
				return style
			})

			// Leave <wps:style> refs unchanged; spPr fill controls the interior color

			// VML fallback (<v:shape>): update fillcolor and convert gradient <v:fill/> to solid
			block = fillColorRe.ReplaceAllStringFunc(block, func(fc string) string {
				return `fillcolor="#` + hex + `"`
			})

			// update VML gradient fill, keep gradient and recolor stops
			block = vFillRe.ReplaceAllStringFunc(block, func(tag string) string {
				getShapeAttr := func(name string) string {
					re := regexp.MustCompile(`(?i)\b` + name + `="([^"]*)"`)
					m := re.FindStringSubmatch(tag)
					if len(m) == 2 {
						return m[1]
					}
					return ""
				}
				rotate := getShapeAttr("rotate")
				angle := getShapeAttr("angle")
				focus := getShapeAttr("focus")

				base := strings.ToLower(strings.TrimPrefix(hex, "#"))
				darker := adjustBrightnessHex(base, 0.25, false)
				lighter := adjustBrightnessHex(base, 0.35, true)

				attrs := []string{`type="gradient"`, `colors="0 #` + darker + `;0.5 #` + base + `;1 #` + lighter + `"`, `color2="#` + lighter + `"`}
				if rotate != "" {
					attrs = append(attrs, `rotate="`+rotate+`"`)
				}
				if angle != "" {
					attrs = append(attrs, `angle="`+angle+`"`)
				}
				if focus != "" {
					attrs = append(attrs, `focus="`+focus+`"`)
				}
				return `<v:fill ` + strings.Join(attrs, " ") + `/>`
			})
		}

		block = placeholderRe.ReplaceAllString(block, "")

		return block
	})
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
