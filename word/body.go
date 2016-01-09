package word

import (
	"encoding/xml"
)

type Body struct {
	XMLName xml.Name `xml:"w:body"`
	Content []interface{}
}
