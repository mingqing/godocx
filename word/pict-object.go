package word

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"image"
	"image/jpeg"
	"image/png"
	"os"

	"github.com/pborman/uuid"
)

type PictObject struct {
	XMLName xml.Name `xml:"w:pict"`
	Shape   *imageShape
	data    []byte
	format  string
	uuid    string
}

type imageShape struct {
	XMLName xml.Name `xml:"v:shape"`
	Style   string   `xml:"style,attr"`
	Data    *imageData
}

type imageData struct {
	XMLName xml.Name `xml:"v:imagedata"`
	Id      string   `xml:"r:id,attr"`
	Title   string   `xml:"o:title,attr"`
}

func newPictObject(data []byte, format string) *PictObject {
	return &PictObject{data: data, format: format}
}
func newPictObject2(id, title string, width, height float64) *PictObject {
	// convert to pt unit
	style := fmt.Sprintf("width:%0.2fpt;height:%0.2fpt", width/1.333333, height/1.333333)

	shape := &imageShape{Style: style,
		Data: &imageData{Id: id, Title: title}}

	return &PictObject{Shape: shape}
}

func (p *PictObject) Style(width, height float64) {
	style := fmt.Sprintf("width:%0.2fpt;height:%0.2fpt", width/1.333333, height/1.333333)
	shape := &imageShape{Style: style, Data: &imageData{Id: p.GetId(), Title: p.GetTitle()}}
	p.Shape = shape
}

func (p *PictObject) GetImage() (image.Image, error) {
	var err error
	var imgObj image.Image
	switch p.format {
	case ".png":
		imgObj, err = png.Decode(bytes.NewBuffer(p.data))
	case ".jpeg", ".jpg":
		imgObj, err = jpeg.Decode(bytes.NewBuffer(p.data))
	default:
		return nil, errors.New("not support picture")
	}
	if err != nil {
		return nil, err
	}

	return imgObj, nil
}

func (p *PictObject) GetId() string {
	if p.uuid == "" {
		p.uuid = uuid.NewUUID().String()
	}
	return p.uuid
}

func (p *PictObject) GetTitle() string {
	return "image-" + p.GetId()
}

func (p *PictObject) SaveFileTo(imgPath string) error {
	outFile, err := os.Create(imgPath)
	if err != nil {
		return err
	}
	defer outFile.Close()

	b := bufio.NewWriter(outFile)
	if _, err := b.Write(p.data); err != nil {
		return err
	}
	return b.Flush()
}
