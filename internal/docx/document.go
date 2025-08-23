package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"path"
	"regexp"
	"slices"
	"strconv"
	"strings"
	"text/template"

	goziputils "github.com/JJJJJJack/go-zip-utils"
)

type documentMeta struct {
	docPrIdsBijectiveIndex uint32
	docPrIds               []uint32
	greaterCNvPrId         uint64
	greaterRId             uint64
	greaterWP14DocId       uint64
	greaterPictureNumber   uint64
	greaterChartNumber     uint64
}

const DOC_PR_ID_ROOF = 2_147_483_647 // docx id attributes are 32-bit signed integers

// rotl32 rotates a 32-bit integer left by k bits.
func rotl32(x uint32, k uint) uint32 {
	return (x << k) | (x >> (32 - k))
}

// bijective32 is a fast bijective permutation on 32-bit integers.
func bijective32(x uint32) uint32 {
	x *= 0x9E3779B1 // example odd 32-bit multiplier
	x = rotl32(x, 16)
	x ^= 0x85EBCA6B // example 32-bit xor constant
	return x
}

func (d *documentMeta) RandUniqueDocPrId() (uint32, error) {
	if d.docPrIdsBijectiveIndex == 0 {
		d.docPrIdsBijectiveIndex = 1
	}

	docPrId := uint32(0)
	for i := 0; ; i++ {
		if i >= DOC_PR_ID_ROOF {
			return 0, fmt.Errorf("this should not happen, surpassed %d attempts to create a unique id for a wp:docPr tag", DOC_PR_ID_ROOF)
		}

		docPrId = bijective32(d.docPrIdsBijectiveIndex) % DOC_PR_ID_ROOF
		d.docPrIdsBijectiveIndex++

		if !slices.Contains(d.docPrIds, docPrId) && docPrId != 0 {
			break
		}
	}

	d.docPrIds = append(d.docPrIds, docPrId)

	return docPrId, nil
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
func ParseDocumentMeta(zm goziputils.ZipMap) (*documentMeta, error) {
	d := documentMeta{}

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
	mediaList := []MediaRel{}

	re := regexp.MustCompile(`<w:[A-Za-z]?>\[\[IMAGE:.*?\]\]</w:[A-Za-z]>`)
	imagePlaceholderRE := regexp.MustCompile(`\[\[IMAGE:.*?\]\]`)
	xmlBlocks := re.FindAllString(srcXML, -1)
	for _, xmlBlock := range xmlBlocks {
		imageTemplate, err := template.New("image-template").Parse(imageTemplateXml)
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
		docPrId, err := d.RandUniqueDocPrId()
		if err != nil {
			return srcXML, mediaList, fmt.Errorf("unable to get unique docPrId: %w", err)
		}

		rid := d.NextRId()
		rId := fmt.Sprintf("rId%d", rid)

		err = imageTemplate.Execute(&buffer, XmlImageData{
			DocPrId: docPrId,
			Name:    filename,
			RefID:   rId,
		})
		if err != nil {
			return srcXML, mediaList, fmt.Errorf("unable to execute image template: %w", err)
		}

		mediaList = append(mediaList, MediaRel{
			Type:   ImageMediaType,
			RefID:  rId,
			Source: path.Join("media", filename),
		})

		srcXML = strings.ReplaceAll(srcXML, xmlBlock, buffer.String())
	}

	return srcXML, mediaList, nil
}

func (d *documentMeta) ApplyTemplate(f *zip.File, zipWriter *zip.Writer, data any) ([]MediaRel, error) {
	documentXml, err := goziputils.ReadZipFileContent(f)
	if err != nil {
		return nil, fmt.Errorf("unable to read document file '%s': %w", f.Name, err)
	}

	documentXml = []byte(PatchXml(string(documentXml)))

	tmpl, err := template.New(f.Name).
		Funcs(template.FuncMap{
			"toImage": toImage,
		}).
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

	output = postProcessing(output)

	err = goziputils.RewriteFileIntoZipWriter(zipWriter, f, []byte(output))
	if err != nil {
		return nil, fmt.Errorf("unable to rewrite file '%s' in zip: %w", f.Name, err)
	}

	return media, nil
}
