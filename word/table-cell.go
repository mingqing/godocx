package word

import (
	"encoding/xml"
)

type TableCell struct {
	XMLName xml.Name `xml:"w:tc"`
	Pr      *CellProps
	Content []interface{}
	P       *Paragraph
}

type CellProps struct {
	XMLName xml.Name `xml:"w:tcPr"`
	W       *CellWidth
	VAlign  *TableVAlign
}

type CellWidth struct {
	XMLName xml.Name `xml:"w:tcW"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}
type CellHeight struct {
	XMLName xml.Name `xml:"w:tcH"`
	W       string   `xml:"w:w,attr,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

func (c *CellProps) Width(width, typ string) {
	c.W = &CellWidth{W: width, Type: typ}
}
func (c *CellProps) Align(val string) {
	c.VAlign = &TableVAlign{Val: val}
}

func NewTableCell() *TableCell {
	return &TableCell{}
}

func (t *TableCell) AddProps() *CellProps {
	t.Pr = &CellProps{}
	return t.Pr
}

func (t *TableCell) AddParagraph() *Paragraph {
	t.P = NewParagraph()
	return t.P
}

func (t *TableCell) Add(obj interface{}) {
	t.Content = append(t.Content, obj)
}
