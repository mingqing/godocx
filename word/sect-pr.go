package word

import (
	"encoding/xml"
)

type sectPr struct {
	XMLName   xml.Name   `xml:"w:sectPr"`
	Val       string     `xml:"w:val,attr,omitempty"` // on,1,true off,0,false
	Header1   *reference `xml:"w:headerReference"`
	Header2   *reference `xml:"w:headerReference"`
	Footer1   *reference `xml:"w:footerReference"`
	Footer2   *reference `xml:"w:footerReference"`
	PgSz      *pgSz
	PgMar     *pgMar
	PgNumType *pgNumType
	Cols      *cols
	DocGrid   *docGrid
}

type reference struct {
	Id   string `xml:"r:id,attr,omitempty"`
	Type string `xml:"r:type,attr,omitempty"`
}

type pgSz struct {
	XMLName xml.Name `xml:"w:pgSz"`
	W       string   `xml:"w:w,attr,omitempty"`
	H       string   `xml:"w:h,attr,omitempty"`
	Code    string   `xml:"w:code,attr,omitempty"`
}

type pgMar struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr,omitempty"`
	Right   string   `xml:"w:right,attr,omitempty"`
	Bottom  string   `xml:"w:bottom,attr,omitempty"`
	Left    string   `xml:"w:left,attr,omitempty"`
	Header  string   `xml:"w:header,attr,omitempty"`
	Footer  string   `xml:"w:footer,attr,omitempty"`
	Gutter  string   `xml:"w:gutter,attr,omitempty"`
}

type pgNumType struct {
	XMLName xml.Name `xml:"w:pgNumType"`
	Start   string   `xml:"w:space"`
}

type cols struct {
	XMLName xml.Name `xml:"w:cols"`
	Space   string   `xml:"w:space,attr"`
}

type docGrid struct {
	XMLName   xml.Name `xml:"w:docGrid"`
	Type      string   `xml:"w:type,attr"`
	LinePitch string   `xml:"w:linePitch,attr"`
}

func newSectPr() *sectPr {
	s := &sectPr{}
	s.Header1 = &reference{Id: "rId6", Type: "even"}
	s.Header2 = &reference{Id: "rId7", Type: "default"}
	s.Footer1 = &reference{Id: "rId8", Type: "even"}
	s.Footer2 = &reference{Id: "rId9", Type: "default"}
	s.PgSz = &pgSz{W: "11906", H: "16838", Code: "9"}
	s.PgMar = &pgMar{Top: "1440", Right: "1797", Bottom: "1440", Left: "1797", Header: "851", Footer: "992", Gutter: "0"}
	s.PgNumType = &pgNumType{Start: "1"}
	s.Cols = &cols{Space: "425"}
	s.DocGrid = &docGrid{Type: "lines", LinePitch: "312"}

	return s
}
