package word

import (
	"encoding/xml"
)

type Document struct {
	XMLName  xml.Name `xml:"w:document"`
	XmlnsO   string   `xml:"xmlns:o,attr,omitempty"`
	XmlnsR   string   `xml:"xmlns:r,attr,omitempty"`
	XmlnsM   string   `xml:"xmlns:m,attr,omitempty"`
	XmlnsV   string   `xml:"xmlns:v,attr,omitempty"`
	XmlnsW   string   `xml:"xmlns:w,attr,omitempty"`
	XmlnsVe  string   `xml:"xmlns:ve,attr,omitempty"`
	XmlnsWp  string   `xml:"xmlns:wp,attr,omitempty"`
	XmlnsW10 string   `xml:"xmlns:w10,attr,omitempty"`
	XmlnsWne string   `xml:"xmlns:wne,attr,omitempty"`
	Body     Body
}

func (d *Document) AddParagraph() *Paragraph {
	paragh := NewParagraph()
	d.Body.Content = append(d.Body.Content, paragh)
	return paragh
}

func (d *Document) AddTable() *Table {
	tbl := NewTable()
	d.Body.Content = append(d.Body.Content, tbl)
	return tbl
}

func NewDocument() *Document {
	d := &Document{}

	return d
}
