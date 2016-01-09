package word

import (
	"encoding/xml"
)

type ComplexScriptFontsize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
