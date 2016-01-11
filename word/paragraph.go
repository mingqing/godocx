package word

import (
	"encoding/xml"
)

type Paragraph struct {
	XMLName      xml.Name `xml:"w:p"`
	RsidR        string   `xml:"w:rsidR,attr,omitempty"`
	RsidP        string   `xml:"w:rsidP,attr,omitempty"`
	RsidRPr      string   `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string   `xml:"w:rsidRDefault,attr,omitempty"`
	Content      []interface{}
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
	r := &RunContent{}
	p.Content = append(p.Content, r)
	return r
}
