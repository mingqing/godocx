package word

import (
	"encoding/xml"
)

type document struct {
	XMLName  xml.Name `xml:"w:document"`
	XmlnsO   string   `xml:"xmlns:o,attr"`
	XmlnsR   string   `xml:"xmlns:r,attr"`
	XmlnsM   string   `xml:"xmlns:m,attr"`
	XmlnsV   string   `xml:"xmlns:v,attr"`
	XmlnsW   string   `xml:"xmlns:w,attr"`
	XmlnsVe  string   `xml:"xmlns:ve,attr"`
	XmlnsWp  string   `xml:"xmlns:wp,attr"`
	XmlnsW10 string   `xml:"xmlns:w10,attr"`
	XmlnsWne string   `xml:"xmlns:wne,attr"`
	//Relationship []relationship `xml:"Relationship"`
}

type docBody struct {
}

func newDocument() *document {
}
