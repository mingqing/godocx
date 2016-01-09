package godocx

import (
	"encoding/xml"
	"fmt"

	"github.com/mingqing/godocx/word"
)

type docXml struct {
}

func NewDocXml() *docXml {
	return &docXml{}
}

func (d *docXml) Save(dirpath string) error {
	c := newContentType()
	return c.Save(dirpath)
}

func (d *docXml) Test() {
	/*
		c := newContentType()
			output, err := xml.MarshalIndent(c, "", "  ")
			if err != nil {
				fmt.Printf("error: %v\n", err)
			}
	*/
	document := word.NewDocument()
	docByte, _ := xml.MarshalIndent(document, "", "  ")

	fmt.Println(xml.Header + string(docByte))
}
