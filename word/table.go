package word

import (
	"encoding/xml"
)

type Table struct {
	XMLName xml.Name `xml:"w:tbl"`
	Pr      *TableProperties
	Grid    *TableGrid
	Tr      []*TableRow
}

func NewTable() *Table {
	return &Table{}
}

func (t *Table) AddProps() *TableProperties {
	t.Pr = &TableProperties{}
	return t.Pr
}

func (t *Table) AddGrid() *TableGrid {
	t.Grid = &TableGrid{}
	return t.Grid
}

func (t *Table) AddRow() *TableRow {
	tr := &TableRow{}
	t.Tr = append(t.Tr, tr)
	return tr
}
