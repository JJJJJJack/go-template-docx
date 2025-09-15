package xml

import (
	"archive/zip"
	"bytes"
	"fmt"

	goziputils "github.com/JJJJJJack/go-zip-utils"
)

// Handler takes the content of a file and returns the modified
// content that will replace it.
type Handler func(content string) (string, error)

// HandlersMap maps filenames to a [Handler] functions chain. Each file content
// will be modified sequentially by each function in the []Handler slice.
// The final output will overwrite the original.
type HandlersMap map[string][]Handler

func ProcessedOutput(filesProcessorsMaps []HandlersMap, outputBuffer *bytes.Buffer, preOrPost string) error {
	for _, filesPostProcessorsMap := range filesProcessorsMaps {
		zipBytes := append([]byte(nil), outputBuffer.Bytes()...)

		zipMap, err := goziputils.NewZipMapFromBytes(zipBytes)
		if err != nil {
			return fmt.Errorf("unable to create zip map during %s-processing: %w", preOrPost, err)
		}

		outputBuffer.Reset()
		outputZipWriter := zip.NewWriter(outputBuffer)

		for filename, f := range zipMap {
			processors := filesPostProcessorsMap[filename]
			if len(processors) == 0 {
				if err := goziputils.CopyFile(outputZipWriter, f); err != nil {
					return fmt.Errorf("unable to copy original file '%s' during %s-processing: %w", f.Name, preOrPost, err)
				}
				continue
			}

			fileContent, err := goziputils.ReadZipFileContent(f)
			if err != nil {
				return fmt.Errorf("unable to read file '%s' during %s-processing: %w", f.Name, preOrPost, err)
			}

			xmlOutput := string(fileContent)
			for _, processor := range processors {
				xmlOutput, err = processor(xmlOutput)
				if err != nil {
					return fmt.Errorf("error %s processing file '%s': %w", preOrPost, f.Name, err)
				}
			}

			if err := goziputils.RewriteFileIntoZipWriter(outputZipWriter, f, []byte(xmlOutput)); err != nil {
				return fmt.Errorf("unable to rewrite %s-processed file '%s': %w", preOrPost, f.Name, err)
			}
		}

		if err := outputZipWriter.Close(); err != nil {
			return fmt.Errorf("unable to close zip writer after %s-processing: %w", preOrPost, err)
		}
	}

	return nil
}
