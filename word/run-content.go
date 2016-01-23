package word

import (
	"encoding/xml"
	"strings"
)

type RunContent struct {
	XMLName xml.Name `xml:"w:r"`
	RsidRPr string   `xml:"w:rsidRPr,attr,omitempty"`
	Content []interface{}
	T       *textContent
	//T       string `xml:"w:t,omitempty"`
	//RPr     *RunProperties `xml:"w:rPr"`
}
type textContent struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Text    string   `xml:",chardata"`
}

func NewRunContent() *RunContent {
	return &RunContent{Content: make([]interface{}, 0)}
}

func (r *RunContent) AddRunProperties() *RunProperties {
	obj := NewRunProperties()
	r.Content = append(r.Content, obj)
	return obj
}

func (r *RunContent) Text(val string) {
	if r.T == nil {
		r.T = &textContent{}
	}

	if strings.TrimSpace(val) == "" {
		r.T.Space = "preserve"
	}
	r.T.Text = val
}

func (r *RunContent) Insert(obj interface{}) {
	r.Content = append(r.Content, obj)
}

func (r *RunContent) AddPict(data []byte, format string) *PictObject {
	pict := newPictObject(data, format)
	r.Content = append(r.Content, pict)
	return pict
}
