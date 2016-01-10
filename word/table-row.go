package word

import (
	"encoding/xml"
)

type TableRow struct {
	XMLName xml.Name `xml:"w:tr"`
	Pr      *RowProps
	Tc      []*TableCell
}

type RowProps struct {
	XMLName xml.Name `xml:"w:trPr"`
	H       *RowHeight
	W       *RowWidth
	Jc      *Align
}
type RowHeight struct {
	XMLName xml.Name `xml:"w:trHeight"`
	Val     string   `xml:"w:val,attr"`
}
type RowWidth struct {
	XMLName xml.Name `xml:"w:trWidth"`
	Val     string   `xml:"w:val,attr"`
}

func (r *RowProps) Align(val string) {
	r.Jc = NewAlign(val)
}

func (r *RowProps) Height(val string) {
	r.H = &RowHeight{Val: val}
}
func (r *RowProps) Width(val string) {
	r.W = &RowWidth{Val: val}
}

func NewTableRow() *TableRow {
	return &TableRow{}
}

func (t *TableRow) AddProps() *RowProps {
	t.Pr = &RowProps{}
	return t.Pr
}

func (t *TableRow) AddCell() *TableCell {
	tc := &TableCell{}
	t.Tc = append(t.Tc, tc)
	return tc
}
