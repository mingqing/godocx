package word

import (
	"encoding/xml"
)

type TableBorders struct {
	XMLName xml.Name `xml:"w:tblBorders"`
	Top     *Border  `xml:"w:top"`
	Left    *Border  `xml:"w:left"`
	Bottom  *Border  `xml:"w:bottom"`
	Right   *Border  `xml:"w:right"`
	InsideH *Border  `xml:"w:insideH"`
	InsideV *Border  `xml:"w:insideV"`
}

type Border struct {
	Color      string `xml:"w:color,attr,omitempty"`
	Frame      string `xml:"w:frame,attr,omitempty"`
	Shadow     string `xml:"w:shadow,attr,omitempty"`
	Space      string `xml:"w:space,attr,omitempty"`
	Sz         string `xml:"w:sz,attr,omitempty"`
	ThemeColor string `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string `xml:"w:themeShade,attr,omitempty"`
	ThemeTint  string `xml:"w:themeTint,attr,omitempty"`
	Val        string `xml:"w:val,attr,omitempty"`
}

func NewTableBorders() *TableBorders {
	return &TableBorders{}
}
