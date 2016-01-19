package word

import (
	"encoding/xml"
)

type RunContent struct {
	XMLName xml.Name `xml:"w:r"`
	RsidRPr string   `xml:"w:rsidRPr,attr,omitempty"`
	Content []interface{}
	T       string `xml:"w:t,omitempty"`
	//RPr     *RunProperties `xml:"w:rPr"`
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
	r.T = val
}

func (r *RunContent) Insert(obj interface{}) {
	r.Content = append(r.Content, obj)
}

func (r *RunContent) AddPict(id, title string, width, height float64) *PictObject {
	pict := newPictObject(id, title, width, height)
	r.Content = append(r.Content, pict)
	return pict
}
