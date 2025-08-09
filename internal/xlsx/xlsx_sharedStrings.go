package xlsx

import (
	"bytes"
	"regexp"
	"strings"
)

// TODO: switch to xml parsing
func GetReferencedSharedStringsByIndexAndCleanup(fileContent []byte) ([]byte, map[int]string) {
	values := make(map[int]string)

	re := regexp.MustCompile(`<si(?:\s[^>]*)?><t(?:\s[^>]*)?>(.*?)</t></si>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	for i, match := range matches {
		splitTag := strings.Split(match[1], "[[NUMBER:")
		if len(splitTag) < 2 {
			continue
		}
		value := strings.Split(splitTag[1], "]]")[0]
		values[i] = value

		fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(""), 1)
	}

	return fileContent, values
}
