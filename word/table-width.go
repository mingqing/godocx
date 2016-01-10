package word

import (
	"encoding/xml"
)

type TableWidth struct {
	XMLName xml.Name `xml:"w:tblW"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

func NewTableWidth(w, typ string) *TableWidth {
	return &TableWidth{W: w, Type: typ}
}
