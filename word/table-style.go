package word

import (
	"encoding/xml"
)

type TableStyle struct {
	XMLName xml.Name `xml:"w:tblStyle"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

func NewTableStyle(val string) *TableStyle {
	return &TableStyle{Val: val}
}
