package word

import (
	"encoding/xml"
)

// Run props
type RunProperties struct {
	XMLName xml.Name `xml:"w:rPr"`
	RFonts  *RunFonts
	B       *Bold
	I       *Italics
	Sz      *Fontsize
	SzCs    *ComplexScriptFontsize
}
