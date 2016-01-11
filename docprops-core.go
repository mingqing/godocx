package godocx

import (
	"encoding/xml"
	"os"
	"path"
	//"fmt"
)

type docPropsCore struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XmlnsCp        string   `xml:"xmlns:cp,attr"`
	XmlnsDc        string   `xml:"xmlns:dc,attr"`
	XmlnsDcterms   string   `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype  string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi       string   `xml:"xmlns:xsi,attr"`
	Title          string   `xml:"dc:title"`
	Subject        string   `xml:"dc:subject"`
	Creator        string   `xml:"dc:creator"`
	Keywords       string   `xml:"cp:keywords"`
	Description    string   `xml:"dc:description"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy"`
	Revision       string   `xml:"cp:revision"`
	Created        string   `xml:"dcterms:created"`
	Modified       string   `xml:"dcterms:modified"`
}

func newDocPropsCore() *docPropsCore {
	c := &docPropsCore{}
	c.XmlnsCp = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	c.XmlnsDc = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

	return c
}

func (c *docPropsCore) Save(dirpath string) error {
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
