package docx

import (
	"regexp"
	"strings"
)

func postProcessing(srcXML string) string {
	// Remove empty table rows
	re := regexp.MustCompile(`<w:tr\b[^>]*>[\s\S]*?</w:tr>`)
	matches := re.FindAllString(srcXML, -1)
	for _, match := range matches {
		if !strings.Contains(match, "<w:t></w:t>") {
			continue
		}
		srcXML = strings.ReplaceAll(srcXML, match, "")
	}

	return srcXML
}
