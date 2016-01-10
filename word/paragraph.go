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
	PPr          *ParagraphProperties
	R            *RunContent
}

func NewParagraph() *Paragraph {
	return &Paragraph{}
}

func (p *Paragraph) AddProperties() *ParagraphProperties {
	p.PPr = NewParagraphProperties()
	return p.PPr
}

func (p *Paragraph) AddRunContent() *RunContent {
	p.R = &RunContent{}
	return p.R
}
