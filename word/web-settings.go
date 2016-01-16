package word

import (
	"encoding/xml"
	"os"
	"path"
)

type WebSettings struct {
	XMLName            xml.Name `xml:"w:webSettings"`
	XmlnsR             string   `xml:"w:xmln:r,attr,omitempty"`
	XmlnsW             string   `xml:"w:xmln:w,attr,omitempty"`
	OptimizeForBrowser string   `xml:"w:optimizeForBrowser"`
	RelyOnVML          string   `xml:"w:relyOnVML"`
	AllowPNG           string   `xml:"w:allowPNG"`
}

func NewWebSettings() *WebSettings {
	sttng := &WebSettings{}
	sttng.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	sttng.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

	return sttng
}

func (n *WebSettings) Save(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(n, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "webSettings.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
