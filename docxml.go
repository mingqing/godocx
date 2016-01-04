package godocx

import (
	"encoding/xml"
	"fmt"
)

/*
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="wmf" ContentType="image/x-wmf"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
*/
type docXml struct {
	XMLName xml.Name `xml:"Types"`
	Xmlns   string   `xml:"xmlns,attr"`
	Default []*typeDefault
}

type typeDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type typeOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

func NewDocXml() *docXml {
	return &docXml{Default: make([]*typeDefault, 0)}
}

func (d *docXml) Test() {
	d.Xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"
	d.Default = append(d.Default, &typeDefault{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"})
	d.Default = append(d.Default, &typeDefault{Extension: "xml", ContentType: "application/xml"})
	d.Default = append(d.Default, &typeDefault{Extension: "png", ContentType: "image/png"})
	d.Default = append(d.Default, &typeDefault{Extension: "wmf", ContentType: "image/x-wmf"})
	d.Default = append(d.Default, &typeDefault{Extension: "bin", ContentType: "application/vnd.openxmlformats-officedocument.oleObject"})

	output, err := xml.MarshalIndent(d, "  ", "    ")
	if err != nil {
		fmt.Printf("error: %v\n", err)
	}

	fmt.Println("str:", string(output))
}
