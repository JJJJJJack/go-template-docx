package xlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// TODO: switch to xml parsing
func UpdateSheet(fileContent []byte, numberCellsValues map[int]string, sharedStringsNewIndexes map[int]int) ([]byte, []string, error) {
	re := regexp.MustCompile(`<c[^>]*t="s"[^>]*>.*?<v>(\d+)</v>.*?</c>`)
	matches := re.FindAllStringSubmatch(string(fileContent), -1)

	valuesOrderedByAppearance := []string{}
	for _, match := range matches {
		sharedStringIndex, err := strconv.Atoi(match[1])
		if err != nil {
			return nil, []string{}, fmt.Errorf("unable to convert index %s to int: %w", match[1], err)
		}

		// put number value directly in sheet cell and remove type attribute
		if numberValue, ok := numberCellsValues[sharedStringIndex]; ok {
			removedRefToSharedString := strings.Replace(match[0], `t="s"`, "", 1)

			oldV := fmt.Sprintf("<v>%d</v>", sharedStringIndex)
			newV := fmt.Sprintf("<v>%s</v>", numberValue)

			replace := strings.Replace(removedRefToSharedString, oldV, newV, 1)

			fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(replace), 1)

			valuesOrderedByAppearance = append(valuesOrderedByAppearance, numberValue)
		}

		// update shared string index if it has changed
		if newSharedStringIndex, ok := sharedStringsNewIndexes[sharedStringIndex]; ok {
			oldV := fmt.Sprintf("<v>%d</v>", sharedStringIndex)
			newV := fmt.Sprintf("<v>%d</v>", newSharedStringIndex)

			replace := strings.Replace(match[0], oldV, newV, 1)

			fileContent = bytes.Replace(fileContent, []byte(match[0]), []byte(replace), 1)
		}
	}

	return fileContent, valuesOrderedByAppearance, nil
}

// Cell represents a <c> element in sheetN.xml
type Cell struct {
	T string `xml:"t,attr"` // type attribute (e.g. "s" for shared string)
}

// Row represents a <row> element
type Row struct {
	Cells []Cell `xml:"c"`
}

// Worksheet represents the <worksheet> root
type Worksheet struct {
	Rows []Row `xml:"sheetData>row"`
}

// GetCountFromXml counts <c t="s"> cells in a sheetN.xml
func GetCountFromXml(data []byte) (uint, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))

	var ws Worksheet
	if err := decoder.Decode(&ws); err != nil {
		return 0, err
	}

	count := uint(0)
	for _, row := range ws.Rows {
		for _, cell := range row.Cells {
			if cell.T == "s" { // shared string cell
				count++
			}
		}
	}

	return count, nil
}
