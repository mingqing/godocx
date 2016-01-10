package word

import (
	"encoding/xml"
)

type TableVAlign struct {
	XMLName xml.Name `xml:"w:vAlign"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

func NewTableVAlign(val string) *TableVAlign {
	return &TableVAlign{Val: val}
}
