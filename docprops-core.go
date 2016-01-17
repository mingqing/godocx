package godocx

import (
	"encoding/xml"
	"os"
	"path"
	"time"
	//"fmt"
)

type docPropsCore struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XmlnsCp        string   `xml:"xmlns:cp,attr"`
	XmlnsDc        string   `xml:"xmlns:dc,attr"`
	XmlnsDcterms   string   `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype  string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi       string   `xml:"xmlns:xsi,attr"`
	Creator        string   `xml:"dc:creator"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy"`
	Revision       string   `xml:"cp:revision"`
	Created        string   `xml:"dcterms:created"`
	Modified       string   `xml:"dcterms:modified"`
	//Title          string   `xml:"dc:title"`
	//Subject        string   `xml:"dc:subject"`
	//Keywords       string `xml:"cp:keywords"`
	//Description    string `xml:"dc:description"`
}

func newDocPropsCore() *docPropsCore {
	c := &docPropsCore{}
	c.XmlnsCp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	c.XmlnsDc = "http://purl.org/dc/elements/1.1/"
	c.XmlnsDcterms = "http://purl.org/dc/terms/"
	c.XmlnsDcmitype = "http://purl.org/dc/dcmitype/"
	c.XmlnsXsi = "http://www.w3.org/2001/XMLSchema-instance"
	c.Creator = "godocx"
	c.LastModifiedBy = "godocx"
	c.Revision = "11"

	nowTimeStr := time.Now().Format(time.RFC3339)
	c.Created = nowTimeStr
	c.Modified = nowTimeStr
	//c.Created = "2016-01-11T07:16:00Z"
	//c.Modified = "2016-01-11T07:16:00Z"

	return c
}

func (c *docPropsCore) Save(dirpath string) error {
	fpath := path.Join(dirpath, "docProps")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(c, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "core.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
