package docx

const (
	ImageMediaType = iota + 1
)

type MediaMap map[string][]byte

type MediaRel struct {
	Type   uint
	RefID  string
	Source string
}
