package docx

import (
	"bytes"
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

func toNumberCell(v any) (interface{}, error) {
	switch v := v.(type) {
	case
		int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return fmt.Sprintf("[[NUMBER:%v]]", v), nil
	}

	fmt.Printf("Type %T not implemented in toNumberCell", v)
	return nil, fmt.Errorf("type %T not implemented in toNumberCell", v)
}

func getSharedStringsValues(fileContent []byte) map[int]string {
	values := make(map[int]string)

	re := regexp.MustCompile(`<t(?:\s[^>]*)?>(.*?)</t>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	for i, match := range matches {
		splitTag := strings.Split(match[1], "[[NUMBER:")
		if len(splitTag) < 2 {
			continue
		}
		value := strings.Split(splitTag[1], "]]")[0]
		values[i] = value
	}

	return values
}

func replaceIndexesWithValuesFromSharedStrings(fileContent []byte, values map[int]string) ([]byte, error) {
	re := regexp.MustCompile(`<c[^>]*t="s"[^>]*>.*?<v>(\d+)</v>.*?</c>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	for i, match := range matches {

		n, err := strconv.Atoi(match[1])
		if err != nil {
			return nil, fmt.Errorf("unable to convert index %s to int: %w", match[1], err)
		}

		if value, ok := values[n]; ok {
			fmt.Println(i, match)
			// fmt.Println("Replacing index", n, "with value", values[n])
			// remove attribute `t="s"`
			// and the value with the value from shared strings
			removedRefToSharedString := strings.Replace(match[0], `t="s"`, "", 1)
			oldV := fmt.Sprintf("<v>%d</v>", n)
			newV := fmt.Sprintf("<v>%s</v>", value)
			fmt.Println(i, "Replacing", oldV, "with", newV)
			replace := strings.Replace(removedRefToSharedString, oldV, newV, 1)
			fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(replace), 1)
		}
	}

	return fileContent, nil
}
