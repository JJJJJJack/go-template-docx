package docx

import (
	"fmt"
	"reflect"
	"regexp"
	"strings"
)

// removeEmptyTableRows removes empty table rows from the provided XML string.
func removeEmptyTableRows(srcXML string) string {
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
