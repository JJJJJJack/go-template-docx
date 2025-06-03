package main

import (
	"fmt"
	"os"
	"testdocx/internal/docx"
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
	templateFile := "report-template.docx"
	templateValues := Data{
		Title:       "Asset Report",
		Description: "Asset Report Description",
		Table: []Row{
			{
				Title:       "Prova1",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "generic.png",
				A:           true,
			},
			{
				Title:       "Prova2",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "ap.png",
				A:           false,
			},
			{
				Title:       "Prova3",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "windows.png",
				A:           true,
			},
		},
		A: "test",
		B: "laa",
		Charts: []ExcelRecord{
			{Label: "Cat1", Value: 111},
			{Label: "Cat2", Value: 222},
			{Label: "Cat3", Value: 333},
			{Label: "Cat4", Value: 444},
		},
		Pie: Pie{
			CriticalSum: 10,
			HighSum:     20,
			MediumSum:   30,
			LowSum:      40,
		},
	}

	template := docx.NewTemplate(templateFile)
	template.Media("generic.png", readFile("generic.png"))
	template.Media("ap.png", readFile("ap.png"))
	template.Media("windows.png", readFile("windows.png"))

	err := template.Apply(templateValues)
	if err != nil {
		fmt.Println("Error applying template:", err)
	}

	err = template.Save("output.docx")
	if err != nil {
		fmt.Println("Error saving template:", err)
	}
}
