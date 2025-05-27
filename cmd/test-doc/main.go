package main

import (
	"os"
	"testdocx/internal/docx"
)

type Row struct {
	Title       string
	Text        string
	Description string
	Icon        string
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
			},
			{
				Title:       "Prova2",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "ap.png",
			},
			{
				Title:       "Prova3",
				Text:        "Lorem Ipsum",
				Description: "Lorem Ipsum descript",
				Icon:        "windows.png",
			},
		},
	}

	template := docx.NewTemplate(templateFile)
	template.Media("generic.png", readFile("generic.png"))
	template.Media("ap.png", readFile("ap.png"))
	template.Media("windows.png", readFile("windows.png"))

	template.Apply(data)
	template.Save("output.docx")
}
