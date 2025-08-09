package xlsx

type XlsxData struct {
	ChartNumbers []string
}
type XlsxFile map[string]XlsxData

// used to store xlsx charts values accross multiple files processings
// instead of argument drilling down the values in each function
// ...any better idea? I hate globals and objects referenced everywhere
var XlsxFiles = make(XlsxFile)
