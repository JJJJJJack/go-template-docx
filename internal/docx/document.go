package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
	"text/template"

	goziputils "github.com/JJJJJJack/go-zip-utils"
)

type documentMeta struct {
	docPrIdsBijectiveIndex uint32
	docPrIds               []uint32
	// greaterCNvPrId         uint64
	greaterRId uint64
	// greaterWP14DocId       uint64
	greaterPictureNumber uint64
	// greaterChartNumber     uint64
	templateFuncs template.FuncMap
}

const DOC_PR_ID_ROOF = 2_147_483_647 // docx id attributes are 32-bit signed integers

// rotl32 rotates a 32-bit integer left by k bits.
func rotl32(x uint32, k uint) uint32 {
	return (x << k) | (x >> (32 - k))
}

// bijective32 is a fast bijective permutation on 32-bit integers.
func bijective32(x uint32) uint32 {
	x *= 0x9E3779B1
	x = rotl32(x, 16)
	x ^= 0x85EBCA6B
	return x
}

func (d *documentMeta) RandUniqueDocPrId() (uint32, error) {
	if d.docPrIdsBijectiveIndex == 0 {
		d.docPrIdsBijectiveIndex = 1
	}

	nextDocPrId := uint32(0)
findNextPrId:
	for i := 0; ; i++ {
		if i >= DOC_PR_ID_ROOF {
			return 0, fmt.Errorf("this should not happen, surpassed %d attempts to create a unique id for a wp:docPr tag", DOC_PR_ID_ROOF)
		}

		nextDocPrId = bijective32(d.docPrIdsBijectiveIndex) % DOC_PR_ID_ROOF
		d.docPrIdsBijectiveIndex++

		for _, docPrId := range d.docPrIds {
			if nextDocPrId == docPrId {
				continue findNextPrId
			}
		}

		if nextDocPrId != 0 {
			break
		}
	}

	d.docPrIds = append(d.docPrIds, nextDocPrId)

	return nextDocPrId, nil
}

func (d *documentMeta) NextPictureNumber() uint64 {
	d.greaterPictureNumber++
	return d.greaterPictureNumber
}

func (d *documentMeta) NextRId() uint64 {
	d.greaterRId++
	return d.greaterRId
}

// TODO: use xml parsing instead of regex
func ParseDocumentMeta(zm goziputils.ZipMap, tf template.FuncMap) (*documentMeta, error) {
	d := documentMeta{
		templateFuncs: template.FuncMap{
			"image":        image,
			"replaceImage": replaceImage,
			// "toCenteredImage": toCenteredImage,
			"preserveNewline":  preserveNewline,
			"breakParagraph":   breakParagraph,
			"shadeTextBg":      shadeTextBg,
			"shapeBgFillColor": shapeBgFillColor,
		},
	}

	for funcName, fn := range tf {
		d.templateFuncs[funcName] = fn
	}

	// work on word/document.xml

	documentFile := zm["word/document.xml"]
	if documentFile == nil {
		return nil, fmt.Errorf("word/document.xml not found in docx")
	}

	documentContent, err := goziputils.ReadZipFileContent(documentFile)
	if err != nil {
		return nil, fmt.Errorf("error reading zip file content: %w", err)
	}

	idAndPictureNRegEx := regexp.MustCompile(`<wp:docPr\s+id="(\d+)"\s+name="Picture\s+(\d+)"\s*/>`)

	docPrAttrsMatches := idAndPictureNRegEx.FindAllStringSubmatch(string(documentContent), -1)
	d.docPrIds = make([]uint32, 0, len(docPrAttrsMatches))
	for _, m := range docPrAttrsMatches {
		docPrId, err := strconv.ParseUint(m[1], 10, 32)
		if err != nil {
			return nil, fmt.Errorf("could not parse DocPr ID '%s': %w", m[1], err)
		}

		d.docPrIds = append(d.docPrIds, uint32(docPrId))

		pictureNumber, err := strconv.ParseUint(m[2], 10, 64)
		if err != nil {
			return nil, fmt.Errorf("could not parse Picture Number '%s': %w", m[2], err)
		}

		if pictureNumber > d.greaterPictureNumber {
			d.greaterPictureNumber = pictureNumber
		}
	}

	// work on word/_rels/document.xml.rels

	wordDocumentRelsFile := zm["word/_rels/document.xml.rels"]
	if wordDocumentRelsFile == nil {
		return nil, fmt.Errorf("word/_rels/document.xml.rels not found in zip")
	}

	wordDocumentRelsContent, err := goziputils.ReadZipFileContent(wordDocumentRelsFile)
	if err != nil {
		return nil, fmt.Errorf("could not read zip file content: %w", err)
	}

	rIdNRegEx := regexp.MustCompile(`"rId(\d+)"`)

	rIdMatches := rIdNRegEx.FindAllStringSubmatch(string(wordDocumentRelsContent), -1)
	for _, match := range rIdMatches {
		num, err := strconv.ParseUint(match[1], 10, 64)
		if err != nil {
			return nil, fmt.Errorf("could not parse rId '%s': %w", match[1], err)
		}

		if num > d.greaterRId {
			d.greaterRId = num
		}
	}

	return &d, nil
}

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
func (d *documentMeta) replaceImages(srcXML string) (string, []MediaRel, error) {
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

	return result, mediaRels, nil
}

// ReplaceAllShapeBgColors finds shapes that contain the [[SHAPE_BG_COLOR:RRGGBB]]/[[SHAPE_BG_COLOR:#RRGGBB]]
// placeholder and uses its value to replace the fillcolor attribute of the shape
// TODO: replace with proper XML parsing
func (d *documentMeta) applyShapesBgFillColor(srcXML string) (string, error) {
	shapeBlockRe := regexp.MustCompile(`(?s)<v:(?:shape|rect|roundrect|oval|line|polyline|arc|curve)\b[^>]*>.*?</v:(?:shape|rect|roundrect|oval|line|polyline|arc|curve)>`)

	placeholderRe := regexp.MustCompile(`\[\[SHAPE_BG_COLOR:#?([0-9A-Fa-f]{6})\]\]`)

	result := shapeBlockRe.ReplaceAllStringFunc(srcXML, func(block string) string {
		pm := placeholderRe.FindStringSubmatch(block)
		if len(pm) < 2 {
			return block
		}

		hex := strings.ToUpper(pm[1])
		if !strings.HasPrefix(hex, "#") {
			hex = "#" + hex
		}

		block = placeholderRe.ReplaceAllString(block, "")

		startTagRe := regexp.MustCompile(`(?s)^<v:(?:roundrect|rect|shape)\b[^>]*>`)
		startTag := startTagRe.FindString(block)
		if startTag == "" {
			return block
		}
		rest := block[len(startTag):]

		fillAttrRe := regexp.MustCompile(`\bfillcolor="[^"]*"`)
		if fillAttrRe.MatchString(startTag) {
			startTag = fillAttrRe.ReplaceAllString(startTag, `fillcolor="`+hex+`"`)
		} else {
			startTag = strings.Replace(startTag, ">", ` fillcolor="`+hex+`">`, 1)
		}

		return startTag + rest
	})

	return result, nil
}

func (d *documentMeta) ApplyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]MediaRel, error) {
	documentXml, err := goziputils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read document file '%s': %w", f.Name, err)
	}

	documentXml = []byte(PatchXml(string(documentXml)))

	tmpl, err := template.New(f.Name).
		Funcs(d.templateFuncs).
		Parse(string(documentXml))
	if err != nil {
		return nil, fmt.Errorf("unable to parse template in file '%s': %w", f.Name, err)
	}

	appliedTemplate := bytes.Buffer{}
	err = tmpl.Execute(&appliedTemplate, data)
	if err != nil {
		return nil, fmt.Errorf("unable to execute template in file '%s': %w", f.Name, err)
	}

	output, media, err := d.applyImages(appliedTemplate.String())
	if err != nil {
		return nil, fmt.Errorf("unable to apply images in file '%s': %w", f.Name, err)
	}

	output, replaceMedia, err := d.replaceImages(output)
	if err != nil {
		return nil, fmt.Errorf("unable to replace images in file '%s': %w", f.Name, err)
	}
	media = append(media, replaceMedia...)

	output, err = d.applyShapesBgFillColor(output)
	if err != nil {
		return nil, fmt.Errorf("unable to apply shapes background fill color in file '%s': %w", f.Name, err)
	}

	output = postProcessing(output)

	err = goziputils.RewriteFileIntoZipWriter(zipWriter, f, []byte(output))
	if err != nil {
		return nil, fmt.Errorf("unable to rewrite file '%s' in zip: %w", f.Name, err)
	}

	return media, nil
}
