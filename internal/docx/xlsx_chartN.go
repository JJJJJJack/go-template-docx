package docx

import (
	"bytes"
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// TODO: switch to xml parsing
func ReplaceSharedStringIndicesWithValues(fileContent []byte, values map[int]string) ([]byte, error) {
	re := regexp.MustCompile(`<c[^>]*t="s"[^>]*>.*?<v>(\d+)</v>.*?</c>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	for _, match := range matches {

		n, err := strconv.Atoi(match[1])
		if err != nil {
			return nil, fmt.Errorf("unable to convert index %s to int: %w", match[1], err)
		}

		if value, ok := values[n]; ok {
			removedRefToSharedString := strings.Replace(match[0], `t="s"`, "", 1)

			oldV := fmt.Sprintf("<v>%d</v>", n)
			newV := fmt.Sprintf("<v>%s</v>", value)

			replace := strings.Replace(removedRefToSharedString, oldV, newV, 1)

			fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(replace), 1)
		}
	}

	return fileContent, nil
}
