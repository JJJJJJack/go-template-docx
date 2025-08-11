package mydebug

import (
	"fmt"
	"strings"
)

func FindAndPrintSnippet(s, substr string) {
	i := strings.Index(s, substr)
	if i >= 0 {
		fmt.Println("found snippet -->", s[i-30:i+30], "<--")
	}
}
