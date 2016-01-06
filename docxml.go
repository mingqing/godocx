package godocx

import (
	"encoding/xml"
	"fmt"
)

type docXml struct {
}

func NewDocXml() *docXml {
	return &docXml{}
}

func (d *docXml) Test() {
	c := newContentType()

	output, err := xml.MarshalIndent(c, "", "  ")
	if err != nil {
		fmt.Printf("error: %v\n", err)
	}

	fmt.Println(xml.Header + string(output))
}
