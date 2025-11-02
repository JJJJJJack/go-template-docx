package docx

import (
	"fmt"
	"regexp"
	"strings"
)

// removeEmptyTableRows removes empty table rows from the provided XML string.
func removeEmptyTableRows(srcXML string) string {
	trRe := regexp.MustCompile(`(?s)<w:tr\b[^>]*>.*?</w:tr>`) // match a table row
	tRe := regexp.MustCompile(`(?is)<w:t\b[^>]*>(.*?)</w:t>`) // capture text content
	visRe := regexp.MustCompile(`(?is)<w:drawing\b|<w:pict\b|<mc:AlternateContent\b|<v:shape\b|<wps:spPr\b`)

	isRowEmpty := func(row string) bool {
		if visRe.MatchString(row) {
			return false
		}

		texts := tRe.FindAllStringSubmatch(row, -1)
		if len(texts) == 0 {
			return true
		}

		for _, m := range texts {
			if strings.TrimSpace(m[1]) != "" {
				return false
			}
		}
		return true
	}

	return trRe.ReplaceAllStringFunc(srcXML, func(row string) string {
		if isRowEmpty(row) {
			return ""
		}
		return row
	})
}

// ensureXmlSpacePreserve ensures all <w:t> elements with leading/trailing
// whitespace have the xml:space="preserve" attribute.
// This is required by Word to preserve spaces; without it, Word collapses whitespace.
func ensureXmlSpacePreserve(srcXML string) string {
	textRe := regexp.MustCompile(`<w:t\b([^>]*)>([\s\S]*?)</w:t>`)
	xmlSpaceRe := regexp.MustCompile(`xml:space="preserve"`)

	return textRe.ReplaceAllStringFunc(srcXML, func(match string) string {
		submatches := textRe.FindStringSubmatch(match)
		if len(submatches) < 3 {
			return match
		}

		attrs := submatches[1]
		text := submatches[2]

		// Skip if already has the attribute or doesn't need it
		hasAttribute := xmlSpaceRe.MatchString(attrs)
		needsAttribute := text != "" && text != strings.TrimSpace(text)

		if hasAttribute || !needsAttribute {
			return match
		}

		return fmt.Sprintf(`<w:t xml:space="preserve">%s</w:t>`, text)
	})
}

// flattenNestedTextRuns fixes cases where a template function that returns
// `<w:rPr>..</w:rPr><w:t>..</w:t>` got injected inside an existing `<w:t>`.
// That produces invalid nesting like:
//
//	<w:t> <w:rPr>..</w:rPr><w:t>text</w:t> </w:t>
//
// Replace it with:
//
//	<w:rPr>..</w:rPr><w:t>text</w:t>
//
// Preserves the xml:space="preserve" attribute if it exists in either tag.
func flattenNestedTextRuns(srcXML string) string {
	nestedRe := regexp.MustCompile(`(?is)<w:t\b([^>]*)>\s*(<w:rPr>[\s\S]*?</w:rPr>)\s*<w:t\b([^>]*)>([\s\S]*?)</w:t>\s*</w:t>`)
	xmlSpaceRe := regexp.MustCompile(`xml:space="preserve"`)

	for nestedRe.MatchString(srcXML) {
		srcXML = nestedRe.ReplaceAllStringFunc(srcXML, func(match string) string {
			submatches := nestedRe.FindStringSubmatch(match)
			if len(submatches) < 5 {
				return match
			}
			outerAttrs := submatches[1]
			rPr := submatches[2]
			innerAttrs := submatches[3]
			text := submatches[4]

			// If either tag had xml:space="preserve", preserve it in the flattened output
			if xmlSpaceRe.MatchString(outerAttrs) || xmlSpaceRe.MatchString(innerAttrs) {
				return fmt.Sprintf(`%s<w:t xml:space="preserve">%s</w:t>`, rPr, text)
			}
			return fmt.Sprintf(`%s<w:t>%s</w:t>`, rPr, text)
		})
	}

	return srcXML
}
