package word

import (
	"encoding/xml"
)

type TableGrid struct {
	XMLName xml.Name `xml:"w:tblGrid"`
	Col     []*GridColumn
}

type GridColumn struct {
	XMLName xml.Name `xml:"w:gridCol"`
	W       string   `xml:"w:w,attr,omitempty"`
}

func NewGridColumn(width string) *GridColumn {
	return &GridColumn{W: width}
}

func (t *TableGrid) AddWidth(width string) *GridColumn {
	col := NewGridColumn(width)
	t.Col = append(t.Col, col)
	return col
}
