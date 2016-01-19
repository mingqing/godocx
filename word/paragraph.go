package word

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"fmt"
	"image/png"
	"io/ioutil"
	"os"

	"github.com/pborman/uuid"
)

type Paragraph struct {
	XMLName      xml.Name `xml:"w:p"`
	RsidR        string   `xml:"w:rsidR,attr,omitempty"`
	RsidP        string   `xml:"w:rsidP,attr,omitempty"`
	RsidRPr      string   `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string   `xml:"w:rsidRDefault,attr,omitempty"`
	Content      []interface{}
	rels         *relationships
	home         string
}

func NewParagraph() *Paragraph {
	return &Paragraph{Content: make([]interface{}, 0)}
}

func (p *Paragraph) AddProperties() *ParagraphProperties {
	ppr := NewParagraphProperties()
	p.Content = append(p.Content, ppr)
	return ppr
}

func (p *Paragraph) AddRunContent() *RunContent {
	r := NewRunContent()
	p.Content = append(p.Content, r)
	return r
}

// 未完成
func (p *Paragraph) AddPictFromFile(fpath string) {
	randomId := uuid.NewUUID()
	id := randomId.String()

	fmt.Println("p home:", p.home)
	rawByte, err := ioutil.ReadFile(fpath)
	if err != nil {
		fmt.Println("read file error:", err)
	}
	img, err := png.Decode(bytes.NewBuffer(rawByte))
	if err != nil {
		fmt.Println("read file error:", err)
	}

	title := "image-" + id

	relObj := relationship{}
	relObj.Id = id
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	relObj.Target = "media/" + title + ".png"
	p.rels.Relationship = append(p.rels.Relationship, relObj)

	imagepath := p.home + "/word/media/" + title + ".png"
	fmt.Println("image path:", imagepath)
	outFile, err := os.Create(imagepath)
	if err != nil {
		fmt.Println("save file error:", err)
	}
	defer outFile.Close()

	b := bufio.NewWriter(outFile)
	png.Encode(b, img)
	b.Flush()

	rt := img.Bounds()
	width := float64(rt.Max.X)
	height := float64(rt.Max.Y)

	r := NewRunContent()
	r.AddPict(id, title, width, height)
	p.Content = append(p.Content, r)
}
