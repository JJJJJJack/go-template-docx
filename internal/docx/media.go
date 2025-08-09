package docx

const (
	ImageMediaType = iota + 1
)

type Media struct {
	Filename string
	Data     []byte
}

type MediaRel struct {
	Type   uint
	RefID  string
	Source string
}
