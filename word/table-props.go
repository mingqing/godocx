package word

import (
	"encoding/xml"
)

type TableProperties struct {
	XMLName xml.Name `xml:"w:tblPr"`
	Style   *TableStyle
	W       *TableWidth
	Jc      *Align
	Borders *TableBorders
	Layout  *TableLayout
	Look    *TableLook
}

func (t *TableProperties) AddStyle(val string) *TableStyle {
	t.Style = &TableStyle{Val: val}
	return t.Style
}

func (t *TableProperties) AddWidth(width, typ string) *TableWidth {
	t.W = &TableWidth{W: width, Type: typ}
	return t.W
}

func (t *TableProperties) AddAlign(val string) *Align {
	t.Jc = NewAlign(val)
	return t.Jc
}

func (t *TableProperties) AddBorders() *TableBorders {
	t.Borders = &TableBorders{}
	/*
		t.Borders.Top = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
		t.Borders.Left = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
		t.Borders.Bottom = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
		t.Borders.Right = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
		t.Borders.InsideH = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
		t.Borders.InsideV = &Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
	*/
	return t.Borders
}

func (t *TableProperties) AddLayout(val string) *TableLayout {
	t.Layout = &TableLayout{Type: val}
	return t.Layout
}

func (t *TableProperties) AddLook(val string) *TableLook {
	t.Look = &TableLook{Val: val}
	return t.Look
}
