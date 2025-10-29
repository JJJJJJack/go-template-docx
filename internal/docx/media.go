package docx

import (
	"bytes"
	"fmt"
	stdimage "image"
	_ "image/jpeg"
	_ "image/png"
	"math"
)

const (
	ImageMediaType = iota + 1
)

const emusPerInch = 914400.0

type Media struct {
	Data         []byte
	WordFilename string
}

type MediaMap map[string]*Media

type MediaRel struct {
	Type   uint
	RefID  string
	Source string
}

func (d *documentMeta) computeDocxImageSize(imageData []byte) (int, int, error) {
	cfg, _, err := stdimage.DecodeConfig(bytes.NewReader(imageData))
	if err != nil {
		return 0, 0, err
	}

	if cfg.Width == 0 || cfg.Height == 0 {
		return 0, 0, fmt.Errorf("invalid image dimensions")
	}

	widthInches := float64(cfg.Width) / 96.0
	heightInches := float64(cfg.Height) / 96.0

	scale := math.Min(d.maxWidthInches/widthInches, d.maxHeightInches/heightInches)
	if scale > 1 {
		scale = 1
	}

	newWidth := widthInches * scale
	newHeight := heightInches * scale

	cx := int(newWidth * emusPerInch)
	cy := int(newHeight * emusPerInch)
	return cx, cy, nil
}
