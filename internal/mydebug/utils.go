package mydebug

import (
	"fmt"
	"strings"
)

func FindAndPrintSnippet(s, substr string) {
	i := strings.Index(s, substr)
	if i >= 0 {
		fmt.Println(s[i : i+60])
	}
}
