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
	Body     docBody  `xml:"w:body"`
}

type docBody struct {
	Test []interface{} `xml:"w:p"`
}

func NewDocument() *document {
	d := &document{}
	/*
		d.XmlnsO = "urn:schemas-microsoft-com:office:office"
		d.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		d.XmlnsM = "http://schemas.openxmlformats.org/officeDocument/2006/math"
		d.XmlnsV = "urn:schemas-microsoft-com:vml"
		d.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		d.XmlnsVe = "http://schemas.openxmlformats.org/markup-compatibility/2006"
		d.XmlnsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		d.XmlnsW10 = "urn:schemas-microsoft-com:office:word"
		d.XmlnsWne = "http://schemas.microsoft.com/office/word/2006/wordml"
	*/

	d.Body = docBody{Test: make([]interface{}, 0)}
	d.Body.Test = append(d.Body.Test, "test1")
	d.Body.Test = append(d.Body.Test, "test2")
	d.Body.Test = append(d.Body.Test, "test3")

	return d
}
