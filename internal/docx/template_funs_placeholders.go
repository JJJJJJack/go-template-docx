package docx

import (
	"bytes"
	"fmt"
	"path"
	"regexp"
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
const withSolidFill = `</a:prstGeom><a:solidFill><a:srgbClr val="ff0000" /></a:solidFill></wps:spPr>`

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

			block = strings.Replace(block, withoutSolidFill, withSolidFill, 1)

			srgbRe := regexp.MustCompile(`(?i)<a:srgbClr\s+val="[^"]*"\s*/>`)
			block = srgbRe.ReplaceAllString(block, `<a:srgbClr val="`+hex+`"/>`)

			fillColorRe := regexp.MustCompile(`(?i)\bfillcolor="[^"]*"`)
			block = fillColorRe.ReplaceAllStringFunc(block, func(fc string) string {
				return `fillcolor="#` + hex + `"`
			})
		}

		block = placeholderRe.ReplaceAllString(block, "")

		return block
	})
}

// replaceTableCellBgColors finds table cells with [[TABLE_CELL_BG_COLOR:#RRGGBB]]
// placeholders, removes the placeholder text (leaving the <w:t> node intact),
// and applies the color to the w:shd/@w:fill attribute.
func (d *documentMeta) replaceTableCellBgColors(srcXML string) string {
	// Regex to match [[SHAPE_BG_FILL_COLOR:#RRGGBB]] in descr or alt
	placeholderRe := regexp.MustCompile(`\[\[SHAPE_BG_FILL_COLOR:#?([0-9A-Fa-f]{6})\]\]`)

	// Match each <mc:AlternateContent> block separately
	altContentRe := regexp.MustCompile(`(?s)<mc:AlternateContent>.*?</mc:AlternateContent>`)

	return altContentRe.ReplaceAllStringFunc(srcXML, func(block string) string {
		placeholders := placeholderRe.FindAllStringSubmatch(block, -1)
		if len(placeholders) == 0 {
			return block
		}

		for _, pm := range placeholders {
			hex := strings.ToLower(pm[1])

			// 1️⃣ Try to replace existing <a:srgbClr val="..."/>
			srgbRe := regexp.MustCompile(`(?i)<a:srgbClr\s+val="[^"]*"\s*/>`)
			if srgbRe.MatchString(block) {
				block = srgbRe.ReplaceAllString(block, `<a:srgbClr val="`+hex+`"/>`)
			} else {
				// 2️⃣ If not found, insert it inside <a:solidFill> or create <a:solidFill>
				solidFillRe := regexp.MustCompile(`(?i)<a:solidFill>`)
				if solidFillRe.MatchString(block) {
					block = solidFillRe.ReplaceAllString(block, `<a:solidFill><a:srgbClr val="`+hex+`"/>`)
				} else {
					// Fallback: create <a:solidFill> at reasonable place in <wps:spPr>
					spPrRe := regexp.MustCompile(`(?s)(<wps:spPr>.*?</wps:spPr>)`)
					block = spPrRe.ReplaceAllStringFunc(block, func(sppr string) string {
						return strings.Replace(sppr, "</wps:spPr>", `<a:solidFill><a:srgbClr val="`+hex+`"/></a:solidFill></wps:spPr>`, 1)
					})
				}
			}

			// Replace fillcolor in VML shapes
			fillColorRe := regexp.MustCompile(`(?i)\bfillcolor="[^"]*"`)
			block = fillColorRe.ReplaceAllStringFunc(block, func(fc string) string {
				return `fillcolor="#` + hex + `"`
			})
		}

		// Remove all placeholders from descr and alt
		block = placeholderRe.ReplaceAllString(block, "")

		return block
	})
}
