package xlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"strings"
)

// SharedStrings is the root <sst>
type SharedStrings struct {
	XMLName     xml.Name `xml:"sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	SI          []SI     `xml:"si"`
}

// SI represents a <si> entry (string item)
type SI struct {
	T string `xml:"t"`
	R []R    `xml:"r"` // optional rich text runs
}

// R represents a <r> rich text run inside <si>
type R struct {
	T string `xml:"t"`
}

var sharedStringsTagsRE = regexp.MustCompile(`<si(?:\s[^>]*)?><t(?:\s[^>]*)?>(.*?)</t></si>`)

// TODO: switch to xml parsing
func GetReferencedSharedStringsByIndexAndCleanup(fileContent []byte) ([]byte, map[int]string, map[int]int, error) {
	// decoder := xml.NewDecoder(bytes.NewReader(fileContent))

	// var sst SharedStrings
	// if err := decoder.Decode(&sst); err != nil {
	// 	return []byte{}, nil, err
	// }

	numberCellsValues := make(map[int]string)
	stringsCellsOldIndexes := make(map[string]int)
	stringsCellsNewIndexes := make(map[int]int)

	matches := sharedStringsTagsRE.FindAllStringSubmatch(string(fileContent), -1)

	for i, match := range matches {
		splitTag := strings.Split(match[1], "[[NUMBER:")
		if len(splitTag) < 2 {
			stringsCellsOldIndexes[match[1]] = i
			continue
		}
		value := strings.Split(splitTag[1], "]]")[0]
		numberCellsValues[i] = value

		fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(""), 1)
	}

	// save the real remaining shared strings with their new indexes
	matches = sharedStringsTagsRE.FindAllStringSubmatch(string(fileContent), -1)
	for i, match := range matches {
		stringsCellsNewIndexes[stringsCellsOldIndexes[match[1]]] = i
	}

	return fileContent, numberCellsValues, stringsCellsNewIndexes, nil
}

// GetUniqueCountFromXML counts the number of <si> tags in sharedStrings.xml
func GetUniqueCountFromXML(data []byte) (int, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var sst SharedStrings
	if err := decoder.Decode(&sst); err != nil {
		return 0, err
	}

	return len(sst.SI), nil
}

// UpdateSharedStringsCounts updates the count and uniqueCount
// attributes in sharedStrings.xml with previously obtained count and recalculated uniqueCount
func UpdateSharedStringsCounts(sharedStringsContent []byte, count uint) ([]byte, error) {
	uniqueCount, err := GetUniqueCountFromXML(sharedStringsContent)
	if err != nil {
		return nil, fmt.Errorf("error counting unique shared strings in sharedStrings.xml: %w", err)
	}

	countAttrRE := regexp.MustCompile(`count="(\d+)"`)
	uniqueCountAttrRE := regexp.MustCompile(`uniqueCount="(\d+)"`)

	sharedStringsContent = countAttrRE.ReplaceAllFunc(sharedStringsContent, func(match []byte) []byte {
		return []byte(fmt.Sprintf(`count="%d"`, count))
	})

	sharedStringsContent = uniqueCountAttrRE.ReplaceAllFunc(sharedStringsContent, func(match []byte) []byte {
		return []byte(fmt.Sprintf(`uniqueCount="%d"`, uniqueCount))
	})

	return sharedStringsContent, nil
}
