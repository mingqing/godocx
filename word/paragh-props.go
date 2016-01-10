package word

import (
	"encoding/xml"
)

// Paragraph Properties
type ParagraphProperties struct {
	XMLName xml.Name `xml:"w:pPr"`
	Align   *Align
	RPr     *RunProperties
}

func NewParagraphProperties() *ParagraphProperties {
	return &ParagraphProperties{}
}

func (p *ParagraphProperties) AddAlign(val string) *Align {
	p.Align = NewAlign(val)
	return p.Align
}

func (p *ParagraphProperties) AddRunProperties() *RunProperties {
	p.RPr = NewRunProperties()
	return p.RPr
}
