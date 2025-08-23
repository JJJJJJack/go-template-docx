[![Go Reference](https://pkg.go.dev/badge/github.com/JJJJJJack/go-template-docx.svg)](https://pkg.go.dev/github.com/JJJJJJack/go-template-docx)

`go get github.com/JJJJJJack/go-template-docx`

# Notes

go-template-docx is based on the golang template standard library, thus it inherits its templating syntax to parse tokens inside the docx file.
The library doesn't change the original files and only reads it into memory to output a new file with the provided template values.

# Usage

First you need to create an instance of the object to load the docx file in and get the high-level APIs, you have 2 options to do so:

```go
DocxTemplate, err := gotemplatedocx.NewDocxTemplateFromBytes(docxBytes)
if err != nil {
  // handle error
}
```

or

```go
DocxTemplate, err := gotemplatedocx.NewDocxTemplateFromFilename(docxFilename)
if err != nil {
  // handle error
}
```

after obtaining the `DocxTemplate` object it exposes the methods to create a new docx file based on the original templated one, let's walk through the usage for each one

> every function is provided with a Godoc comment, you can find all the exposed APIs in the `go_template_docx.go` file

## 1. Loading Media (PNG, JPG only for now...)

```go
myImagePngBytes, _ := os.ReadFile("myimage.png")
DocxTemplate.Media("myimagealias.png", myImagePngBytes)
```

## 2. Applying the template values

> here the `templateValues` variable could be any json marshallable value, the struct fields will be used as keys in the docx to search to access the value

```go
err := DocxTemplate.Apply(templateValues)
if err != nil {
  // handle error
}
```

## 3. Saving the new docx as new file

> for now it overrides an existing file only if the Save method receives a single string argument

```go
err := DocxTemplate.Save(outputFilename)
if err != nil {
  // handle error
}
```

## 4. Read back bytes from new docx

```go
output := DocxTemplate.Bytes()
```

Enjoy programmatically templating docx files from golang!

# Docx template instructions examples

Let's say we have this json value as our `templateValues` variable:

```json
{
  "Title": "Asset Report",
  "Description": "Asset Report Description",
  "Table": [
    {
      "Title": "Try1",
      "Text": "Text1",
      "Icon": "computer.png"
    },
    {
      "Title": "Try2",
      "Text": "Text2",
      "Icon": "ap.png"
    },
    {
      "Title": "Try3",
      "Text": "Text3",
      "Icon": "windows.png"
    }
  ],
  "ClustCol": [
    { "Label": "Cat1", "Value": 111.11 },
    { "Label": "Cat2", "Value": 222 },
    { "Label": "Cat3", "Value": 333.33 },
    { "Label": "Cat4", "Value": 444 }
  ],
  "Pie": {
    "CriticalSum": 10,
    "HighSum": 20,
    "MediumSum": 30,
    "LowSum": 40
  }
}
```
> these value types covers most of the common use cases

and this is the templated docx file that we load into the `DocxTemplate`:
![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/docx-template-example.png)

with respectively the `Series 1` "Clustered Column" chart

![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/series-1-chart.png)

and the `Vulnerabilities` "Pie" chart

![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/vulnerabilities-chart.png)

now if we run this code
```go
computerPng, err := os.ReadFile(computerPngFilepath)
apPng, err := os.ReadFile(apPngFilepath)
windowsPng, err := os.ReadFile(windowsPngFilepath)

template, _ := gotemplatedocx.NewDocxTemplateFromFilename(docxFilename)

template.Media("computer.png", computerPng)
template.Media("ap.png", apPng)
template.Media("windows.png", windowsPng)

template.Apply(templateValues)

template.Save("output.docx")
```

the `output.docx` file will be the result of the templating engine:

![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/output.docx.png)

now let's walk a into each of the template instructions used in the docx file...


### 1. Fields replacement
- `{{.Title}}` -> `Asset Report`
- `{{.Description}}` -> `Asset Report Description`
> note that the color of the title field is not inherited by its value

### 2. Conditional fields 
- `{{if eq .A "test"}}true{{else}}false{{end}}` -> is an example of logical comparison, checks if field `.A` is equal to the string `"test"`, placing `true` or `false` based on `.A` value
> note that the color of the text is not lost after the conditional expression is parsed

### 3. Iterating arrays
- `{{range .Table}}` -> iterates over the `Table` field which contains an array, for each item in the array it replaces the fields inside the loop
- `{{.Title}}` -> `Try1`, `Try2`, `Try3`
- `{{.Text}}` -> `Text1`, `Text2`, `Text3`
- `{{toImage .Icon}}` -> looks for the media filenames `"computer.png"`, `"ap.png"`, `"windows.png"` loaded through `template.Media(...)` and puts media reference in place

### 4. Indexing array items (Series 1 chart)
- `{{(index .ClustCol 0).Label}}` -> `Cat1`
- `{{ toNumberCell (index .ClustCol 0).Value}}` -> `111.11`
- `{{(index .ClustCol 1).Label}}` -> `Cat2`
- `{{ toNumberCell (index .ClustCol 1).Value}}` -> `222`
- `{{(index .ClustCol 2).Label}}` -> `Cat3`
- `{{ toNumberCell (index .ClustCol 2).Value}}` -> `333.33`
- `{{(index .ClustCol 3).Label}}` -> `Cat4`
- `{{ toNumberCell (index .ClustCol 3).Value}}` -> `444`

there are 3 important things to notice here:
1. the floating point numbers are set to be displayed as 2 digits precision directly from the chart "format cells..." tool
2. the `index` operator takes 2 parameters `.ClustCol` (the field that must contain an array) and the index of the item you want to access, I wrap it all with round parenthesis and after that I can access the fields of the indexed item
3. the `toNumberCell` function is used to format the number in a way that it can be parsed by the docx file as a number, this is useful to properly create functioning charts

### 5. Accessing fields in a nested structure (Pie chart)
- `{{ toNumberCell .Pie.CriticalSum}}` -> `10`
- `{{ toNumberCell .Pie.HighSum}}` -> `20`
- `{{ toNumberCell .Pie.MediumSum}}` -> `30`
- `{{ toNumberCell .Pie.LowSum}}` -> `40`
> still note that we use the `toNumberCell` function to set the cell type and make it readable by the docx chart
