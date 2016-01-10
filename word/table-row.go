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
	Jc      *Align
}

func (r *RowProps) Align(val string) {
	r.Jc = NewAlign(val)
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
