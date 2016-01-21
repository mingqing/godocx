package word

import (
	"encoding/base64"
	"encoding/xml"
	"io/ioutil"
	"path"
	"strings"
)

type Paragraph struct {
	XMLName      xml.Name `xml:"w:p"`
	RsidR        string   `xml:"w:rsidR,attr,omitempty"`
	RsidP        string   `xml:"w:rsidP,attr,omitempty"`
	RsidRPr      string   `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string   `xml:"w:rsidRDefault,attr,omitempty"`
	Content      []interface{}
	rels         *relationships
	home         string
}

func NewParagraph() *Paragraph {
	return &Paragraph{Content: make([]interface{}, 0)}
}

func (p *Paragraph) AddProperties() *ParagraphProperties {
	ppr := NewParagraphProperties()
	p.Content = append(p.Content, ppr)
	return ppr
}

func (p *Paragraph) AddRunContent() *RunContent {
	r := NewRunContent()
	p.Content = append(p.Content, r)
	return r
}

func (p *Paragraph) AddPictFromFile(fpath string) (*PictObject, error) {
	rawByte, err := ioutil.ReadFile(fpath)
	if err != nil {
		return nil, err
	}

	imgExt := strings.ToLower(path.Ext(fpath))

	r := NewRunContent()
	pict := r.AddPict(rawByte, imgExt)
	imgObj, err := pict.GetImage()
	if err != nil {
		return nil, err
	}

	imgTitle := pict.GetTitle()
	relObj := relationship{Id: pict.GetId(),
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: "media/" + imgTitle + imgExt}
	p.rels.Relationship = append(p.rels.Relationship, relObj)

	imgPath := p.home + "/word/media/" + imgTitle + imgExt
	pict.SaveFileTo(imgPath)

	rt := imgObj.Bounds()
	width := float64(rt.Max.X)
	height := float64(rt.Max.Y)
	pict.Style(width, height)

	p.Content = append(p.Content, r)
	return pict, nil
}

func (p *Paragraph) AddPictFromBase64(content, format string, width, height float64) (*PictObject, error) {
	rawByte, err := base64.StdEncoding.DecodeString(content)
	if err != nil {
		return nil, err
	}

	imgExt := "." + format

	r := NewRunContent()
	pict := r.AddPict(rawByte, imgExt)
	imgObj, err := pict.GetImage()
	if err != nil {
		return nil, err
	}

	imgTitle := pict.GetTitle()
	relObj := relationship{Id: pict.GetId(),
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: "media/" + imgTitle + imgExt}
	p.rels.Relationship = append(p.rels.Relationship, relObj)

	imgPath := p.home + "/word/media/" + imgTitle + imgExt
	pict.SaveFileTo(imgPath)

	if width == 0 && height == 0 {
		rt := imgObj.Bounds()
		width = float64(rt.Max.X)
		height = float64(rt.Max.Y)
	}

	pict.Style(width, height)

	p.Content = append(p.Content, r)
	return pict, nil
}
