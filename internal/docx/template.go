package docx

import (
	"regexp"
	"strings"
)

// PatchXml removes automatically insert content between template expressions
// (EG: "{{ .Text }}" could have correctors highlights tags separating the expressions tokens).
func PatchXml(srcXml string) string {
	// Fix separated {{
	re := regexp.MustCompile(`\{([^\}]*?)\{`)
	srcXml = re.ReplaceAllString(srcXml, "{{")

	// Fix separated }}
	re = regexp.MustCompile(`\}([^\{]*?)\}`)
	srcXml = re.ReplaceAllString(srcXml, "}}")

	// Remove unecessary XML tags
	re = regexp.MustCompile(`\{\{[\s\S]*?\}\}`)
	matches := re.FindAllString(srcXml, -1)
	for _, match := range matches {
		xmlRegex := regexp.MustCompile(`(<\s*\/?[\w-:.]+(\s+[^>]*?)?[\s\/]*>)`)
		templateText := xmlRegex.ReplaceAllString(match, "")
		srcXml = strings.ReplaceAll(srcXml, match, templateText)
	}

	return srcXml
}
