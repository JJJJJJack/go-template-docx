package main

import (
	"encoding/json"
	"fmt"
	"os"

	gotemplatedocx "github.com/JJJJJJack/go-template-docx"
)

type ExcelRecord struct {
	Label string
	Value any
}

type Row struct {
	Title       string
	Text        string
	Description string
	Icon        string
	A           bool
}

type Pie struct {
	CriticalSum uint
	HighSum     uint
	MediumSum   uint
	LowSum      uint
}

type Data struct {
	Charts      []ExcelRecord
	Title       string
	Description string
	Table       []Row
	A           string
	B           string
	Pie         Pie
}

func readFile(filename string) []byte {
	data, _ := os.ReadFile(filename)
	return data
}

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: go run main.go <document.docx> <template_values.json>")
		return
	}

	docxFilename := os.Args[1]
	jsonFilename := os.Args[2]

	jsonBytes, err := os.ReadFile(jsonFilename)
	if err != nil {
		fmt.Println("Error reading JSON file:", err)
		return
	}

	templateValues := any(nil)
	err = json.Unmarshal(jsonBytes, &templateValues)
	if err != nil {
		fmt.Println("Error unmarshalling JSON:", err)
		return
	}

	docxBytes := readFile(docxFilename)

	template, err := gotemplatedocx.NewDocxTemplateFromBytes(docxBytes)
	if err != nil {
		fmt.Println("Error creating template:", err)
		return
	}

	// template, err := gotemplatedocx.NewDocxTemplateFromFilename(docxFilename)
	// if err != nil {
	// 	fmt.Println("Error creating template:", err)
	// 	return
	// }

	template.Media("generic.png", readFile("generic.png"))
	template.Media("ap.png", readFile("ap.png"))
	template.Media("windows.png", readFile("windows.png"))

	err = template.Apply(templateValues)
	if err != nil {
		fmt.Println("Error applying template:", err)
	}

	err = template.Save()
	if err != nil {
		fmt.Println("Error saving template:", err)
	}
}
