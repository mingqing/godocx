package word

import (
	"encoding/xml"
	"os"
	"path"
)

type Endnotes struct {
	XMLName  xml.Name `xml:"w:endnotes"`
	XmlnsVe  string   `xml:"xmlns:ve,attr,omitempty"`
	XmlnsO   string   `xml:"xmlns:o,attr,omitempty"`
	XmlnsR   string   `xml:"xmlns:r,attr,omitempty"`
	XmlnsM   string   `xml:"xmlns:m,attr,omitempty"`
	XmlnsV   string   `xml:"xmlns:v,attr,omitempty"`
	XmlnsWp  string   `xml:"xmlns:wp,attr,omitempty"`
	XmlnsW10 string   `xml:"xmlns:w10,attr,omitempty"`
	XmlnsW   string   `xml:"xmlns:w,attr,omitempty"`
	XmlnsWne string   `xml:"xmlns:wne,attr,omitempty"`
	Content  []interface{}
}

type endnote struct {
	XMLName xml.Name `xml:"w:endnote"`
	Type    string   `xml:"w:type,attr,omitempty"`
	Id      string   `xml:"w:id,attr,omitempty"`
	P       *Paragraph
}

func NewEndnotes() *Endnotes {
	notes := &Endnotes{}
	notes.XmlnsVe = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	notes.XmlnsO = "urn:schemas-microsoft-com:office:office"
	notes.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	notes.XmlnsM = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	notes.XmlnsV = "urn:schemas-microsoft-com:vml"
	notes.XmlnsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	notes.XmlnsW10 = "urn:schemas-microsoft-com:office:word"
	notes.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	notes.XmlnsWne = "http://schemas.microsoft.com/office/word/2006/wordml"

	p1 := NewParagraph()
	r1 := p1.AddRunContent()
	r1.Insert(&separator{})
	node1 := &endnote{Type: "separator", Id: "-1", P: p1}
	notes.Content = append(notes.Content, node1)

	p2 := NewParagraph()
	r2 := p2.AddRunContent()
	r2.Insert(&continuationSeparator{})
	node2 := &endnote{Type: "continuationSeparator", Id: "0", P: p2}
	notes.Content = append(notes.Content, node2)

	return notes
}

func (n *Endnotes) Save(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(n, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "endnotes.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
