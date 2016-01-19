package word

import (
	"encoding/xml"
	"fmt"
)

type PictObject struct {
	XMLName xml.Name `xml:"w:pict"`
	Shape   *imageShape
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

func newPictObject(id, title string, width, height float64) *PictObject {
	// convert to pt unit
	style := fmt.Sprintf("width:%0.2fpt;height:%0.2fpt", width/1.333333, height/1.333333)

	shape := &imageShape{Style: style,
		Data: &imageData{Id: id, Title: title}}

	return &PictObject{Shape: shape}
}
