package godocx

import (
	"encoding/xml"
)

type relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationship []relationship `xml:"Relationship"`
}

type relationship struct {
	Id     string `xml:"Id"`
	Type   string `xml:"Type"`
	Target string `xml:"Target"`
}

func newRelationships() *relationships {
	c := &relationships{Relationship: make([]relationship, 0)}
	c.Xmlns = "http://schemas.openxmlformats.org/package/2006/relationships"

	relObj := relationship{}
	relObj.Id = "rId1"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	relObj.Target = "word/document.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId2"
	relObj.Type = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	relObj.Target = "docProps/core.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId3"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	relObj.Target = "docProps/app.xml"
	c.Relationship = append(c.Relationship, relObj)
	return c
}
