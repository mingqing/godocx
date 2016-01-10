package word

import (
	"encoding/xml"
)

type TableLook struct {
	XMLName xml.Name `xml:"w:tblLook"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

func NewTableLook(val string) *TableLook {
	return &TableLook{Val: val}
}
