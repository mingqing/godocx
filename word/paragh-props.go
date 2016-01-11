package word

import (
	"encoding/xml"
)

// Paragraph Properties
type ParagraphProperties struct {
	XMLName xml.Name `xml:"w:pPr"`
	Align   *Align
	RPr     *RunProperties
	Content []interface{}
}

func NewParagraphProperties() *ParagraphProperties {
	return &ParagraphProperties{Content: make([]interface{}, 0)}
}

func (p *ParagraphProperties) AddAlign(val string) *Align {
	align := NewAlign(val)
	p.Content = append(p.Content, align)
	return align
}

func (p *ParagraphProperties) AddRunProperties() *RunProperties {
	rpr := NewRunProperties()
	p.Content = append(p.Content, rpr)
	return rpr
}
