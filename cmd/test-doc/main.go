package main

import (
	"fmt"
	"os"
	"testdocx/internal/docx"
)

type Row struct {
	Title       string
	Text        string
	Description string
	Icon        string
	A           bool
}

type Data struct {
	Chart struct {
		Category1 string
		Category2 string
		Category3 string
		Category4 string
	}
	Title       string
	Description string
	Table       []Row
	A           string
	B           string
}

func readFile(filename string) []byte {
	data, _ := os.ReadFile(filename)
	return data
}

func main() {
	templateFile := "report-template.docx"
	data := Data{
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
		Chart: struct {
			Category1 string
			Category2 string
			Category3 string
			Category4 string
		}{
			Category1: "Category 1",
			Category2: "Category 2",
			Category3: "Category 3",
			Category4: "Category 4",
		},
	}

	template := docx.NewTemplate(templateFile)
	template.Media("generic.png", readFile("generic.png"))
	template.Media("ap.png", readFile("ap.png"))
	template.Media("windows.png", readFile("windows.png"))

	err := template.Apply(data)
	if err != nil {
		fmt.Println("Error applying template:", err)
	}

	err = template.Save("output.docx")
	if err != nil {
		fmt.Println("Error saving template:", err)
	}
}
