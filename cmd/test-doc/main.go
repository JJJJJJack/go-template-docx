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
	A           string
}

type Data struct {
	Title string
	Table []Row
}

func readFile(filename string) []byte {
	data, _ := os.ReadFile(filename)
	return data
}

func main() {
	templateFile := "report-template.docx"
	data := Data{
		Title: "Asset Report",
		Table: []Row{
			{
				Title:       "Prova1",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "generic.png",
				A:           "test",
			},
			{
				Title:       "Prova2",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "ap.png",
				A:           "test2",
			},
			{
				Title:       "Prova3",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "windows.png",
				A:           "test3",
			},
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
