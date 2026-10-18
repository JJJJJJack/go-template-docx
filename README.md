[![Go Reference](https://pkg.go.dev/badge/github.com/JJJJJJack/go-template-docx.svg)](https://pkg.go.dev/github.com/JJJJJJack/go-template-docx)

[documentation](https://jjjjjjack.github.io/go-template-docx/getting-started).

`go get github.com/JJJJJJack/go-template-docx`

⭐ Please consider supporting by starring the repo ⭐

# Notes

```diff
! the `“/"` problem: if you have issues with using template expressions with constant values passed inline check that you are not using the default double quotes for word `“` or `”` instead of using the ascii double quote `"`, this breaks the golang template library making it unable to take arguments
```

go-template-docx is based on the golang template standard library, thus it inherits its templating syntax to parse tokens inside the docx file.
The library doesn't change the original files and only reads it into memory to output a new file with the provided template values.

- supports go1.18+
- based on the golang template library syntaxes with features such as:
- supports adding your own custom template functions
- supports text styling
- supports images (png|jpg)
- supports embedded charts templating
- supports tables templating
- supports shapes
- supports preserving text formatting (color, bold, italic, font size, etc...) when replacing text
- [more...](https://jjjjjjack.github.io/go-template-docx/getting-started)
