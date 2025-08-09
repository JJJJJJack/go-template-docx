package xlsx

import "fmt"

func toNumberCell(v any) (interface{}, error) {
	switch v := v.(type) {
	case
		int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return fmt.Sprintf("[[NUMBER:%v]]", v), nil
	}

	return nil, fmt.Errorf("type %T not implemented in toNumberCell", v)
}
