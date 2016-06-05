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
	rels     *relationships
	home     string
}

func NewDocument(home string) *Document {
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

	d.rels = newRelationships()
	d.home = home

	fpath := path.Join(home, "word")
	os.Mkdir(fpath, os.ModePerm)

	os.Mkdir(path.Join(fpath, "_rels"), os.ModePerm)
	os.Mkdir(path.Join(fpath, "media"), os.ModePerm)

	return d
}

func (d *Document) AddParagraph() *Paragraph {
	paragh := NewParagraph()
	paragh.rels = d.rels
	paragh.home = d.home
	d.Body.Content = append(d.Body.Content, paragh)
	return paragh
}
func (d *Document) NewParagraph() *Paragraph {
	paragh := NewParagraph()
	paragh.rels = d.rels
	paragh.home = d.home
	return paragh
}

func (d *Document) AddTable() *Table {
	tbl := NewTable()
	d.Body.Content = append(d.Body.Content, tbl)
	return tbl
}

func (d *Document) SectPr(size string, binding, answer bool) {
	if size == "b4" {
		d.Body.Content = append(d.Body.Content, newSectPrB4(binding, answer))
	} else if size == "b5" {
		d.Body.Content = append(d.Body.Content, newSectPrB5(binding, answer))
	} else {
		d.Body.Content = append(d.Body.Content, newSectPr(binding, answer))
	}
}
func (d *Document) MiddleSectPr(size string, binding, answer bool) {
	p := d.AddParagraph()
	prp := p.AddProperties()
	if size == "b4" {
		prp.Content = append(prp.Content, newSectPrB4(binding, answer))
	} else if size == "b5" {
		prp.Content = append(prp.Content, newSectPrB5(binding, answer))
	} else {
		prp.Content = append(prp.Content, newSectPr(binding, answer))
	}
}

func (d *Document) Save(dirpath string) error {
	if err := d.rels.save(dirpath); err != nil {
		return err
	}

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	os.Mkdir(path.Join(fpath, "_rels"), os.ModePerm)
	os.Mkdir(path.Join(fpath, "media"), os.ModePerm)

	// sectPr
	//d.Body.Content = append(d.Body.Content, newSectPr())

	output, err := xml.MarshalIndent(d, "", "")
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
