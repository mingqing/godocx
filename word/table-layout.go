package word

import (
	"encoding/xml"
)

type TableLayout struct {
	XMLName xml.Name `xml:"w:tblLayout"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

func NewTableLayout(typ string) *TableLayout {
	return &TableLayout{Type: typ}
}
