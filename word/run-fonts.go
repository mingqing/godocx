package word

import (
	"encoding/xml"
)

type RunFonts struct {
	XMLName       xml.Name `xml:"w:rFonts"`
	Ascii         string   `xml:"w:ascii,attr,omitempty"`
	EastAsia      string   `xml:"w:eastAsia,attr,omitempty"`
	AsciiTheme    string   `xml:"w:asciiTheme,attr,omitempty"`
	EastAsiaTheme string   `xml:"w:eastAsiaTheme,attr,omitempty"`
	HansiTheme    string   `xml:"w:hAnsiTheme,attr,omitempty"`
	Hint          string   `xml:"w:hint,attr,omitempty"`
}
