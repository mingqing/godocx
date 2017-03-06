package word

import (
	"encoding/xml"
)

type TableInd struct {
	XMLName xml.Name `xml:"w:tblInd"`
	Val     string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

func NewTableInd(val, typ string) *TableInd {
	return &TableInd{Val: val, Type: typ}
}
