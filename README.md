`go get github.com/JJJJJJack/go-template-docx`

# Notes

go-template-docx is based on the golang template standard library, thus it inherits its templating syntax to parse tokens inside the docx file.
The library doesn't change the original files and only reads it into memory to output a new file with the provided template values.

# Usage

First you need to create an instance of the object to load the docx file in and exposes the high-level APIs, you have 2 options to do it:

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

# 1. Loading Media (PNG, JPG only for now...)
```go
myImagePngBytes, _ := os.ReadFile("myimage.png")
docxTemplate.Media("myimagealias.png", myImagePngBytes)
```

# 2. Applying the template values
> here the `templateValues` variable could be any json marshallable value, the struct fields will be used as keys in the docx to search to access the value
```go
err := docxTemplate.Apply(templateValues)
if err != nil {
  // handle error
}
```

# 3. Saving the new docx as new file
> for now it overrides an existing file only if the Save method receives a single string argument
```go
err := docxTemplate.Save(outputFilename)
if err != nil {
  // handle error
}
```

# 4. Read back bytes from new docx
```go
output := docxTemplate.Bytes()
```

Enjoy programmatically templating docx files from golang!
