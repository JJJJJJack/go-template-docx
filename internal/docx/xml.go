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

	// Remove unnecessary XML tags inside template expressions and unescape XML entities
	re = regexp.MustCompile(`\{\{[\s\S]*?\}\}`)
	matches := re.FindAllString(srcXml, -1)
	for _, match := range matches {
		xmlRegex := regexp.MustCompile(`(<\s*\/?[\w-:.]+(\s+[^>]*?)?[\s\/]*>)`)
		templateText := xmlRegex.ReplaceAllString(match, "")
		// Unescape common XML/HTML entities that Word may inject inside attributes
		// (e.g., {{shapeBgFillColor (index .Map &quot;Color2&quot;)}})
		templateText = strings.NewReplacer(
			"&quot;", "\"",
			"&#34;", "\"",
			"&apos;", "'",
			"&#39;", "'",
			"&lt;", "<",
			"&#60;", "<",
			"&gt;", ">",
			"&#62;", ">",
			// &amp; MUST be last to avoid double-unescaping
			"&amp;", "&",
			"&#38;", "&",
		).Replace(templateText)

		srcXml = strings.ReplaceAll(srcXml, match, templateText)
	}

	// Word may strip quotes inside certain attribute values (e.g., alt/descr of shapes).
	// That leads to invalid Go template syntax like: {{shapeBgFillColor 00FF00}}.
	// To make templating robust, wrap bare hex arguments in quotes for known funcs.
	// Examples fixed here:
	//   - {{shapeBgFillColor 00FF00}}   -> {{shapeBgFillColor "00FF00"}}
	//   - {{shapeBgFillColor #00FF00}}  -> {{shapeBgFillColor "#00FF00"}}
	//   - {{tableCellBgColor 00FF00}}   -> {{tableCellBgColor "00FF00"}}
	wrapBareHexArg := func(funcName string) {
		pat := regexp.MustCompile(`(?i)\{\{\s*` + funcName + `\s+(#?[0-9A-Fa-f]{6})\s*\}\}`)
		srcXml = pat.ReplaceAllString(srcXml, `{{`+funcName+` "$1"}}`)
	}

	wrapBareHexArg("shapeBgFillColor")
	wrapBareHexArg("tableCellBgColor")

	return srcXml
}
