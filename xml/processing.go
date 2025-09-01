package xml

// Handler takes the content of a file and returns the modified
// content that will replace it.
type Handler func(content string) (string, error)

// HandlersMap maps filenames to a [Handler] functions chain. Each file content
// will be modified sequentially by each function in the []Handler slice.
// The final output will overwrite the original.
type HandlersMap map[string][]Handler
