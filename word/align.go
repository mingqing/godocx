package word

import (
	"encoding/xml"
)

type Align struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

func NewAlign(val string) *Align {
	return &Align{Val: val}
}
