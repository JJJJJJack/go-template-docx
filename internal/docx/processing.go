package docx

import (
	"fmt"
	"reflect"
	"regexp"
	"strings"
)

// removeEmptyTableRows removes empty table rows from the provided XML string.
func removeEmptyTableRows(srcXML string) string {
    // A row should be considered "empty" only if it has:
    // - no non‑whitespace text inside any <w:t>...</w:t>
    // - and no visual content (drawings/shapes/alternate content blocks)
    // Word frequently emits empty <w:t></w:t> runs as layout artifacts; removing
    // any row that merely contains one of those is too aggressive and drops
    // legitimate rows. This refined check keeps rows that visibly render.

    trRe := regexp.MustCompile(`(?s)<w:tr\b[^>]*>.*?</w:tr>`) // match a table row
    tRe := regexp.MustCompile(`(?is)<w:t[^>]*>(.*?)</w:t>`)     // capture text content
    visRe := regexp.MustCompile(`(?is)<w:drawing\b|<w:pict\b|<mc:AlternateContent\b|<v:shape\b|<wps:spPr\b`)

    isRowEmpty := func(row string) bool {
        if visRe.MatchString(row) {
            return false
        }

        texts := tRe.FindAllStringSubmatch(row, -1)
        if len(texts) == 0 {
            // No text nodes and no visual content => treat as empty
            return true
        }

        for _, m := range texts {
            if strings.TrimSpace(m[1]) != "" { // any visible char keeps the row
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

// flattenNestedTextRuns fixes cases where a template function that returns
// `<w:rPr>..</w:rPr><w:t>..</w:t>` got injected inside an existing `<w:t>`.
// That produces invalid nesting like:
//   <w:t> <w:rPr>..</w:rPr><w:t>text</w:t> </w:t>
// Replace it with:
//   <w:rPr>..</w:rPr><w:t>text</w:t>
func flattenNestedTextRuns(srcXML string) string {
    nestedRe := regexp.MustCompile(`(?is)<w:t[^>]*>\s*(<w:rPr>[\s\S]*?</w:rPr>)\s*<w:t[^>]*>([\s\S]*?)</w:t>\s*</w:t>`) // greedy across whitespace
    for nestedRe.MatchString(srcXML) {
        srcXML = nestedRe.ReplaceAllString(srcXML, `${1}<w:t>${2}</w:t>`)
    }
    return srcXML
}

// guardSpaces wraps the input text in
// <w:t xml:space="preserve">...</w:t> if it has
// leading or trailing spaces.
func guardSpaces(text string) string {
	if strings.TrimSpace(text) == text {
		return text
	}

	return fmt.Sprintf(`<w:t xml:space="preserve">%s</w:t>`, text)
}

// preserveWhitespaces recursively processes data with reflection,
// to check for strings with leading or trailing spaces
// and preserve them.
func preserveWhitespaces(data any) any {
	if data == nil {
		return nil
	}

	rv := reflect.ValueOf(data)

	switch rv.Kind() {
	case reflect.Ptr:
		if rv.IsNil() {
			return data
		}
		elem := preserveWhitespaces(rv.Elem().Interface())
		ptr := reflect.New(rv.Type().Elem())
		ptr.Elem().Set(reflect.ValueOf(elem))
		return ptr.Interface()

	case reflect.Interface:
		if rv.IsNil() {
			return data
		}
		return preserveWhitespaces(rv.Elem().Interface())

	case reflect.Struct:
		out := reflect.New(rv.Type()).Elem()
		for i := 0; i < rv.NumField(); i++ {
			field := rv.Field(i)
			if !field.CanInterface() || !out.Field(i).CanSet() {
				continue
			}
			processed := preserveWhitespaces(field.Interface())
			val := reflect.ValueOf(processed)
			if val.Type().AssignableTo(out.Field(i).Type()) {
				out.Field(i).Set(val)
			} else if val.Type().ConvertibleTo(out.Field(i).Type()) {
				out.Field(i).Set(val.Convert(out.Field(i).Type()))
			}
		}
		return out.Interface()

	case reflect.Slice:
		if rv.IsNil() {
			return data
		}
		out := reflect.MakeSlice(rv.Type(), rv.Len(), rv.Len())
		for i := 0; i < rv.Len(); i++ {
			processed := preserveWhitespaces(rv.Index(i).Interface())
			val := reflect.ValueOf(processed)
			if val.IsValid() {
				if val.Type().AssignableTo(rv.Type().Elem()) {
					out.Index(i).Set(val)
				} else if val.Type().ConvertibleTo(rv.Type().Elem()) {
					out.Index(i).Set(val.Convert(rv.Type().Elem()))
				}
			}
		}
		return out.Interface()

	case reflect.Array:
		out := reflect.New(rv.Type()).Elem()
		for i := 0; i < rv.Len(); i++ {
			processed := preserveWhitespaces(rv.Index(i).Interface())
			val := reflect.ValueOf(processed)
			if val.IsValid() {
				if val.Type().AssignableTo(rv.Type().Elem()) {
					out.Index(i).Set(val)
				} else if val.Type().ConvertibleTo(rv.Type().Elem()) {
					out.Index(i).Set(val.Convert(rv.Type().Elem()))
				}
			}
		}
		return out.Interface()

	case reflect.Map:
		if rv.IsNil() {
			return data
		}
		out := reflect.MakeMap(rv.Type())
		for _, key := range rv.MapKeys() {
			processed := preserveWhitespaces(rv.MapIndex(key).Interface())
			val := reflect.ValueOf(processed)
			if val.IsValid() {
				if val.Type().AssignableTo(rv.Type().Elem()) {
					out.SetMapIndex(key, val)
				} else if val.Type().ConvertibleTo(rv.Type().Elem()) {
					out.SetMapIndex(key, val.Convert(rv.Type().Elem()))
				}
			}
		}
		return out.Interface()

	case reflect.String:
		s := rv.String()
		if !strings.HasPrefix(s, " ") && !strings.HasSuffix(s, " ") {
			return s
		}

		return guardSpaces(s)

	default:
		return data
	}
}
