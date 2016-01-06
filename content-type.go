package godocx

import (
	"encoding/xml"
	//"fmt"
)

/*
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="wmf" ContentType="image/x-wmf"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/word/headereven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/headerdefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footereven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footerdefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/headeranswer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footeranswer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>
*/
type contentType struct {
	XMLName  xml.Name       `xml:"Types"`
	Xmlns    string         `xml:"xmlns,attr"`
	Default  []typeDefault  `xml:"Default"`
	Override []typeOverride `xml:"Override"`
}

type typeDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type typeOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

func newContentType() *contentType {
	c := &contentType{Default: make([]typeDefault, 0),
		Override: make([]typeOverride, 0)}
	c.Xmlns = "http://schemas.openxmlformats.org/package/2006/content-types"

	defaultValues := make(map[string]string, 0)
	defaultValues["rels"] = "application/vnd.openxmlformats-package.relationships+xml"
	defaultValues["xml"] = "application/xml"
	defaultValues["png"] = "image/png"
	defaultValues["wmf"] = "image/x-wmf"
	defaultValues["bin"] = "application/vnd.openxmlformats-officedocument.oleObject"

	defaultObj := typeDefault{}
	for key, value := range defaultValues {
		defaultObj.Extension = key
		defaultObj.ContentType = value
		c.Default = append(c.Default, defaultObj)
	}

	overrideValues := make(map[string]string, 0)
	overrideValues["/word/document.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

	overrideObj := typeOverride{}
	for key, value := range overrideValues {
		overrideObj.PartName = key
		overrideObj.ContentType = value
		c.Override = append(c.Override, overrideObj)
	}

	return c
}
