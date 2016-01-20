package word

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"image"
	"image/jpeg"
	"image/png"
	"io/ioutil"
	"os"
	"path"
	"strings"
	//"fmt"

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
func (p *Paragraph) AddPictFromFile(fpath string) (*PictObject, error) {
	rawByte, err := ioutil.ReadFile(fpath)
	if err != nil {
		return nil, err
	}

	var imgObj image.Image
	imgExt := strings.ToLower(path.Ext(fpath))
	switch imgExt {
	case ".png":
		imgObj, err = png.Decode(bytes.NewBuffer(rawByte))
	case ".jpeg", ".jpg":
		imgObj, err = jpeg.Decode(bytes.NewBuffer(rawByte))
	default:
		return nil, errors.New("not support picture")
	}
	if err != nil {
		return nil, err
	}

	imgId := uuid.NewUUID().String()
	imgTitle := "image-" + imgId
	relObj := relationship{}
	relObj.Id = imgId
	relObj.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	relObj.Target = "media/" + imgTitle + imgExt
	p.rels.Relationship = append(p.rels.Relationship, relObj)

	imgPath := p.home + "/word/media/" + imgTitle + imgExt
	outFile, err := os.Create(imgPath)
	if err != nil {
		return nil, err
	}
	defer outFile.Close()

	b := bufio.NewWriter(outFile)
	png.Encode(b, imgObj)
	b.Flush()

	rt := imgObj.Bounds()
	width := float64(rt.Max.X)
	height := float64(rt.Max.Y)

	r := NewRunContent()
	pict := r.AddPict(imgId, imgTitle, width, height)
	p.Content = append(p.Content, r)

	return pict, nil
}

func (p *Paragraph) AddPictFromBase64(content, format string) (*PictObject, error) {
	return nil, nil
}
