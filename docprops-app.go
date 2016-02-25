package godocx

import (
	"encoding/xml"
	"os"
	"path"
	//"fmt"
)

type docPropsApp struct {
	XMLName              xml.Name `xml:"Properties"`
	Xmlns                string   `xml:"xmlns,attr"`
	XmlnsVt              string   `xml:"xmlns:vt,attr"`
	Template             string   `xml:"Template"`
	TotalTime            int      `xml:"TotalTime"`
	Pages                int      `xml:"Pages"`
	Words                int      `xml:"Words"`
	Characters           int      `xml:"Characters"`
	Application          string   `xml:"Application"`
	DocSecurity          int      `xml:"DocSecurity"`
	Lines                int      `xml:"Lines"`
	Paragraphs           int      `xml:"Paragraphs"`
	ScaleCrop            bool     `xml:"ScaleCrop"`
	Company              string   `xml:"Company"`
	LinksUpToDate        bool     `xml:"LinksUpToDate"`
	CharactersWithSpaces int      `xml:"CharactersWithSpaces"`
	SharedDoc            bool     `xml:"SharedDoc"`
	HyperlinksChanged    bool     `xml:"HyperlinksChanged"`
	AppVersion           string   `xml:"AppVersion"`
}

func newDocPropsApp() *docPropsApp {
	c := &docPropsApp{}
	c.Xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	c.XmlnsVt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
	c.Template = "Normal.dotm"
	c.TotalTime = 1
	c.Pages = 1
	c.Words = 2
	c.Characters = 30
	c.Application = "Microsoft Office Word"
	c.DocSecurity = 0
	c.Lines = 1
	c.Paragraphs = 1
	c.ScaleCrop = false
	c.Company = "godocx"
	c.LinksUpToDate = false
	c.CharactersWithSpaces = 32
	c.SharedDoc = false
	c.HyperlinksChanged = false
	c.AppVersion = "12.000"

	return c
}

func (c *docPropsApp) Save(dirpath string) error {
	fpath := path.Join(dirpath, "docProps")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(c, "", "")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "app.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
