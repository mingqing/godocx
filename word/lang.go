package word

import (
	"encoding/xml"
)

type lang struct {
	XMLName  xml.Name `xml:"w:lang"`
	Val      string   `xml:"w:val,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	Bidi     string   `xml:"w:bidi,attr,omitempty"`
}

func NewLang() *lang {
	return &lang{}
}
