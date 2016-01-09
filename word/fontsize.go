package word

import (
	"encoding/xml"
)

type Fontsize struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr,omitempty"`
}
