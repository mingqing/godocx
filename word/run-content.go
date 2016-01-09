package word

import (
	"encoding/xml"
)

type RunContent struct {
	XMLName xml.Name       `xml:"w:r"`
	RsidRPr string         `xml:"w:rsidRPr,attr,omitempty"`
	RPr     *RunProperties `xml:"w:rPr"`
	T       string         `xml:"w:t"`
}