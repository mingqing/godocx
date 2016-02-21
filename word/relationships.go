package word

import (
	"encoding/xml"
	"os"
	"path"
)

type relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Xmlns        string         `xml:"xmlns,attr"`
	Relationship []relationship `xml:"Relationship"`
}

type relationship struct {
	Id     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

func newRelationships() *relationships {
	c := &relationships{Relationship: make([]relationship, 0)}
	c.Xmlns = "http://schemas.openxmlformats.org/package/2006/relationships"

	relObj := relationship{}
	relObj.Id = "rId1"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	relObj.Target = "styles.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId2"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	relObj.Target = "settings.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId3"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	relObj.Target = "webSettings.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId4"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	relObj.Target = "footnotes.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId5"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	relObj.Target = "endnotes.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId6"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	relObj.Target = "header1.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId7"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	relObj.Target = "header2.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId8"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	relObj.Target = "footer1.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId9"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	relObj.Target = "footer2.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId10"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	relObj.Target = "fontTable.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId11"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	relObj.Target = "theme/theme1.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId12"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	relObj.Target = "header3.xml"
	c.Relationship = append(c.Relationship, relObj)

	relObj.Id = "rId13"
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	relObj.Target = "header4.xml"
	c.Relationship = append(c.Relationship, relObj)

	return c
}

func (c *relationships) save(dirpath string) error {
	fpath := path.Join(dirpath, "word", "_rels")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(c, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "document.xml.rels"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
