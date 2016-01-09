package word

import (
	"encoding/xml"
)

type Bold struct {
	XMLName xml.Name `xml:"w:b"`
	Val     string   `xml:"w:val,attr,omitempty"` // on,1,true off,0,false
}
