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

			srgbRe := regexp.MustCompile(`(?i)<a:srgbClr\s+val="[^"]*"\s*/>`)
			if !srgbRe.MatchString(block) {
				block = strings.Replace(block, withoutSolidFill, withSolidFill, 1)
			}

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
