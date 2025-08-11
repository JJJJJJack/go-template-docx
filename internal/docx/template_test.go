package docx

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
	"testing"

	"github.com/JJJJJJack/go-template-docx/internal/mydebug"
)

func obtainDocumentXml(filepath string) (string, error) {
	r, err := zip.OpenReader(filepath)
	if err != nil {
		return "", fmt.Errorf("unable to open template file %s: %w", filepath, err)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name != "word/document.xml" {
			continue
		}

		documentFile, err := f.Open()
		if err != nil {
			return "", fmt.Errorf("unable to open document file %s: %w", f.Name, err)
		}

		documentXML, err := io.ReadAll(documentFile)
		if err != nil {
			return "", fmt.Errorf("unable to read document file %s: %w", f.Name, err)
		}

		return string(documentXML), nil
	}

	return "", fmt.Errorf("document.xml not found in the template")
}

func Test_patchXML(t *testing.T) {
	type args struct {
		srcXML string
	}
	type testArgs struct {
		name   string
		srcXml string
		want   string
	}

	// PREPARE TESTS
	tests := []testArgs{}
	for _, filepath := range []string{
		// "../../report-template-libreoffice.docx",
		// "../../report-template-word.docx",
		// "../../report-template.docx",
		"../../output.docx",
	} {
		documentXml, err := obtainDocumentXml(filepath)
		if err != nil {
			t.Fatalf("Failed to obtain document.xml string: %v", err)
		}

		tests = append(tests, testArgs{
			name:   "Search for the if-else condition inside " + filepath,
			srcXml: documentXml,
			want:   "trueyes",
		})
	}

	// ACTUAL TEST
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := PatchXml(tt.srcXml)
			mydebug.FindAndPrintSnippet(got, "true")

			if false == strings.Contains(got, tt.want) {
				t.Errorf("patchXML() result does not contain want '%s'", tt.want)
			}
		})
	}
}
