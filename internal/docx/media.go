package docx

const (
	ImageMediaType = iota + 1
)

type media struct {
	filename string
	data     []byte
}

type mediaRel struct {
	Type   uint
	RefID  string
	Source string
}
