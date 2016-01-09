package word

import (
	"encoding/xml"
)

type Italics struct {
	XMLName xml.Name `xml:"w:i"`
	Val     string   `xml:"w:val,attr,omitempty"` // on,1,true off,0,false
}
