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

func (c *contentType) Save(dirpath string) error {
	output, err := xml.MarshalIndent(c, "", "  ")
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
