package word

import (
	"encoding/xml"
	"os"
	"path"
)

type Document struct {
	XMLName  xml.Name `xml:"w:document"`
	XmlnsVe  string   `xml:"xmlns:ve,attr,omitempty"`
	XmlnsO   string   `xml:"xmlns:o,attr,omitempty"`
	XmlnsR   string   `xml:"xmlns:r,attr,omitempty"`
	XmlnsM   string   `xml:"xmlns:m,attr,omitempty"`
	XmlnsV   string   `xml:"xmlns:v,attr,omitempty"`
	XmlnsWp  string   `xml:"xmlns:wp,attr,omitempty"`
	XmlnsW10 string   `xml:"xmlns:w10,attr,omitempty"`
	XmlnsW   string   `xml:"xmlns:w,attr,omitempty"`
	XmlnsWne string   `xml:"xmlns:wne,attr,omitempty"`
	Body     Body
}

func NewDocument() *Document {
	d := &Document{}

	d.XmlnsO = "urn:schemas-microsoft-com:office:office"
	d.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	d.XmlnsM = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	d.XmlnsV = "urn:schemas-microsoft-com:vml"
	d.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	d.XmlnsVe = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	d.XmlnsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	d.XmlnsW10 = "urn:schemas-microsoft-com:office:word"
	d.XmlnsWne = "http://schemas.microsoft.com/office/word/2006/wordml"

	return d
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

func (d *Document) Save(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(d, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "document.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
