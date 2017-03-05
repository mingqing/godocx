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

// Spacing
type pPrSpacing struct {
	XMLName  xml.Name `xml:"w:spacing"`
	Line     string   `xml:"w:line,attr,omitempty"`
	LineRule string   `xml:"w:lineRule,attr,omitempty"`
}

// numPr
type wNumPr struct {
	XMLName xml.Name    `xml:"w:numPr"`
	Ilvl    *settingVal `xml:"w:ilvl"`
	NumId   *settingVal `xml:"w:numId"`
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

func (p *ParagraphProperties) AddSpacing(line string) *ParagraphProperties {
	t := &pPrSpacing{Line: line, LineRule: "auto"}
	p.Content = append(p.Content, t)
	return p
}

func (p *ParagraphProperties) AddNumPr(ilvl, numId string) *wNumPr {
	t := &wNumPr{Ilvl: &settingVal{Val: ilvl}, NumId: &settingVal{Val: numId}}
	p.Content = append(p.Content, t)
	return t
}
