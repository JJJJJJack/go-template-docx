[![Go Reference](https://pkg.go.dev/badge/github.com/JJJJJJack/go-template-docx.svg)](https://pkg.go.dev/github.com/JJJJJJack/go-template-docx)

`go get github.com/JJJJJJack/go-template-docx`

# Notes

```diff
! if you have issues with using template expressions with constant values passed inline check that you are not using the default double quotes for word `“` or `”` instead of using the ascii double quote `"`, this breaks the golang template library making it not unable to take arguments
```

go-template-docx is based on the golang template standard library, thus it inherits its templating syntax to parse tokens inside the docx file.
The library doesn't change the original files and only reads it into memory to output a new file with the provided template values.
> I'll make a good documentation website asap

- supports go1.18+
- based on the golang template library syntaxes with features such as:
  - text replacement
  - loops
  - conditional statements
  - array indexing
  - nested structures/arrays
  - supports adding your own custom template functions
- supports text styling
- supports images (png|jpg)
  - the `image` template function places an image in the docx file
  - the `replaceImage` template function replaces an existing image in the docx file, useful to keep the image size, position, style and other properties
- supports embedded charts templating:
  - the `toNumberCell` template function sets the chart cell a readable number to make a graphically evaluable chart
- supports tables templating
  - the `tableCellBgColor` template function changes the table cell background fill color by putting the template expression in the cell text
- supports shapes
  - the `shapeBgFillColor` template function changes the shape background fill color by putting the template expression in the shape's alt-text
- supports preserving text formatting (color, bold, italic, font size, etc...) when replacing text

# Template functions list

- `image(filename string)`: the filename parameter looks for an equal loaded `Media`'s filename
> `{{image .ImageFilename}}`
- `replaceImage(filename string)`: the filename parameter looks for an equal loaded `Media`'s filename, it replaces the image inside a `<w:drawing>...</w:drawing>` block, useful to keep the image size and position
> `{{replaceImage .ImageFilename}}` inside the `alt-text` of the image to replace
- `preserveNewline(text string)`: newlines are treated as `SHIFT + ENTER` input, thus keeping the text in the same paragraph.
> `{{preserveNewline .TextWithNewlines}}`
- `breakParagraph(s string)`: newlines are treated as `ENTER` input, thus creating a new paragraph for the sequent line.
> `{{breakParagraph .TextWithNewlines}}`
- `shadeTextBg(hex string, s string)`: applies a background color to the given text, hex string must be in the format `RRGGBB` or `#RRGGBB`
> `{{shadeTextBg .TextBgHex .Text}}`
- `shapeBgFillColor(hex string)`: changes the shape's background fill color, hex string must be in the format `RRGGBB` or `#RRGGBB`
> `{{shapeBgFillColor .ShapeBgHex}}` inside the shape's alt-text
- `toNumberCell(v any)`: (for excel sheets, like charts) sets the cell type to number, useful to make charts work properly, v can be any type that can be converted to a float64
> `{{toNumberCell .Number}}` inside the cell text
- `tableCellBgColor(hex string)`: changes the table cell background fill color, hex string must be in the format `RRGGBB` or `#RRGGBB`
> `{{tableCellBgColor .TableCellBgHex}}` inside the table cell text
- `styledText(s ...string)`: applies multiple styles to the given text, the first argument is the text, the following arguments are styles, supported (case insensitive) styles are:
  - `b` or `bold`
  - `i` or `italic`
  - `u` or `underline`
  - `s` or `strikethrough`
  - `fontsize:<size>` | `fs:<size>` where `<size>` is the font size in points (e.g. for 12pt font size use `fontsize:12` or `fs:12`)
> `{{styledText .Text "b" "i" "fs:14"}}` to apply bold, italic and 14pt font size to the text, this is the most efficient way to apply multiple styles at once
- `bold(s string)`
> `{{bold .Text}}`
- `italic(s string)`
> `{{italic .Text}}`
- `underline(s string)`
> `{{underline .Text}}`
- `strike(s string)`
> `{{strike .Text}}`
- `fontSize(s string, size int)`
> `{{fontSize .Text 14}}` to apply 14pt font size to the text

# Usage

First you need to create an instance of the object to load the docx file in and get the high-level APIs, you have 2 options to do so:

```go
docxTemplate, err := gotemplatedocx.NewDocxTemplateFromBytes(docxBytes)
if err != nil {
  // handle error
}
```

or

```go
docxTemplate, err := gotemplatedocx.NewDocxTemplateFromFilename(docxFilename)
if err != nil {
  // handle error
}
```

after obtaining the `docxTemplate` object it exposes the methods to create a new docx file based on the original templated one, let's walk through the usage for each one

> every function is provided with a Godoc comment, you can find all the exposed APIs in the `go_template_docx.go` file

## 1. Loading Media (PNG, JPG only for now...)

```go
myImagePngBytes, _ := os.ReadFile("myimage.png")
docxTemplate.Media("myimagealias.png", myImagePngBytes)
```

## 2. Adding your custom template functions
```go
docxTemplate.AddTemplateFuncs("appendHeart", func(s string) string {
  return s + " <3"
})
```
> now you can use `{{appendHeart .Text}}` in the docx template to append a heart to the value of `Text`, note that this is one of many possible function prototypes that template.FuncMap supports, full doc on https://pkg.go.dev/text/template#FuncMap

## 3. Applying the template values

> here the `templateValues` variable could be any json marshallable value, the struct fields will be used as keys in the docx to search to access the value

```go
err := docxTemplate.Apply(templateValues)
if err != nil {
  // handle error
}
```

## 4. Saving the new docx as new file

```go
err := docxTemplate.Save(outputFilename)
if err != nil {
  // handle error
}
```

## 5. Read back bytes from new docx

```go
output := docxTemplate.Bytes()
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

or its Go struct equivalent:

```go
type TableRow struct {
		Title string
		Text  string
		Icon  string
}
type ClustColRow struct {
  Label string
  Value float64
}
type PieChartData struct {
  CriticalSum int
  HighSum     int
  MediumSum   int
  LowSum      int
}
type TemplateValues struct {
  A           string // returns errors if is used in template but not defined here
  Title       string
  Description string
  Table       []TableRow
  ClustCol    []ClustColRow
  Pie         PieChartData
}

templateValues = TemplateValues{
  A:           "test2",
  Title:       "Asset Report",
  Description: "Asset Report Description",
  Table: []TableRow{
    {
      Title: "Try1",
      Text:  "Text1",
      Icon:  "computer.png",
    },
    {
      Title: "Try2",
      Text:  "Text2",
      Icon:  "ap.png",
    },
    {
      Title: "Try3",
      Text:  "Text3",
      Icon:  "windows.png",
    },
  },
  ClustCol: []ClustColRow{
    {Label: "Cat1", Value: 111.11},
    {Label: "Cat2", Value: 222},
    {Label: "Cat3", Value: 333.33},
    {Label: "Cat4", Value: 444},
  },
  Pie: PieChartData{
    CriticalSum: 10,
    HighSum:     20,
    MediumSum:   30,
    LowSum:      40,
  },
}
```
> both are evaluated in the same way by the template, except for the `A` field that needs to be present in the struct to not return errors in the template, these value types covers most of the common use cases

and this is the templated docx file that we load into the `docxTemplate`:
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
- `{{image .Icon}}` -> looks for the media filenames `"computer.png"`, `"ap.png"`, `"windows.png"` loaded through `template.Media(...)` and puts media reference in place

### 4. Indexing array items (Series 1 chart)
- `{{(index .ClustCol 0).Label}}` -> `Cat1`
- `{{toNumberCell (index .ClustCol 0).Value}}` -> `111.11`
- `{{(index .ClustCol 1).Label}}` -> `Cat2`
- `{{toNumberCell (index .ClustCol 1).Value}}` -> `222`
- `{{(index .ClustCol 2).Label}}` -> `Cat3`
- `{{toNumberCell (index .ClustCol 2).Value}}` -> `333.33`
- `{{(index .ClustCol 3).Label}}` -> `Cat4`
- `{{toNumberCell (index .ClustCol 3).Value}}` -> `444`

there are 3 important things to notice here:
1. the floating point numbers are set to be displayed as 2 digits precision directly from the chart "format cells..." tool
2. the `index` operator takes 2 parameters `.ClustCol` (the field that must contain an array) and the index of the item you want to access, I wrap it all with round parenthesis and after that I can access the fields of the indexed item
3. the `toNumberCell` function is used to format the number in a way that it can be parsed by the docx file as a number, this is useful to properly create functioning charts

### 5. Accessing fields in a nested structure (Pie chart)
- `{{toNumberCell .Pie.CriticalSum}}` -> `10`
- `{{toNumberCell .Pie.HighSum}}` -> `20`
- `{{toNumberCell .Pie.MediumSum}}` -> `30`
- `{{toNumberCell .Pie.LowSum}}` -> `40`
> still note that we use the `toNumberCell` function to set the cell type and make it readable by the docx chart

### 6. Replacing an image while preserving its style properties
- `{{replaceImage .ImageFilename}}` -> looks for the media filename loaded through `template.Media(...)` equal to the `ImageFilename` field

To use the `replaceImage` function you need to:
1. insert an image in the docx file where you want to place the new image
![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/replaceimage1.png)
2. right click on the image and select "Edit Alt Text"
![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/replaceimage2.png)
3. write the template expression `{{replaceImage .ImageFilename}}`
![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/replaceimage3.png)
4. now when you run the templating engine the image will be replaced while preserving its size, position and other properties
![](https://github.com/JJJJJJack/jubilant-fortnight/blob/main/go-template-docx/replaceimage4.png)
