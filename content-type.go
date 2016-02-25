package godocx

import (
	"encoding/xml"
	"os"
	"path"
	//"fmt"
)

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

	defaultObj := typeDefault{}
	for key, value := range defaultValues {
		defaultObj.Extension = key
		defaultObj.ContentType = value
		c.Default = append(c.Default, defaultObj)
	}

	overrideValues := make(map[string]string, 0)
	overrideValues["/word/document.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	overrideValues["/word/styles.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	overrideValues["/word/endnotes.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
	overrideValues["/docProps/app.xml"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	overrideValues["/word/settings.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	overrideValues["/word/footer2.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	overrideValues["/word/footer1.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	overrideValues["/word/footnotes.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
	overrideValues["/word/theme/theme1.xml"] = "application/vnd.openxmlformats-officedocument.theme+xml"
	overrideValues["/word/header2.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/word/fontTable.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
	overrideValues["/word/webSettings.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
	overrideValues["/word/header1.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/docProps/core.xml"] = "application/vnd.openxmlformats-package.core-properties+xml"
	overrideValues["/word/header3.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/word/header4.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/word/header5.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/word/header6.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	overrideValues["/word/footer3.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	overrideValues["/word/footer4.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

	overrideObj := typeOverride{}
	for key, value := range overrideValues {
		overrideObj.PartName = key
		overrideObj.ContentType = value
		c.Override = append(c.Override, overrideObj)
	}

	return c
}

func (c *contentType) Save(dirpath string) error {
	output, err := xml.MarshalIndent(c, "", "")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(dirpath, "[Content_Types].xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
