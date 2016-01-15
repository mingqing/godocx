package word

import (
	"encoding/xml"
	"os"
	"path"
)

type FontTable struct {
	XMLName xml.Name `xml:"w:fonts"`
	XmlnsR  string   `xml:"w:xmln:r,attr,omitempty"`
	XmlnsW  string   `xml:"w:xmln:w,attr,omitempty"`
	Content []interface{}
}

type font struct {
	XMLName xml.Name   `xml:"w:font"`
	Name    string     `xml:"w:name,attr"`
	AltName *fontValue `xml:"w:altName"`
	Panose1 *fontValue `xml:"w:panose1"`
	Charset *fontValue `xml:"w:charset"`
	Family  *fontValue `xml:"w:family"`
	Pitch   *fontValue `xml:"w:pitch"`
	Sig     *fontSig
}

type fontValue struct {
	Val string `xml:"w:val,attr"`
}
type fontSig struct {
	XMLName xml.Name `xml:"w:sig"`
	Usb0    string   `xml:"w:usb0,attr,omitempty"`
	Usb1    string   `xml:"w:usb1,attr,omitempty"`
	Usb2    string   `xml:"w:usb2,attr,omitempty"`
	Usb3    string   `xml:"w:usb3,attr,omitempty"`
	Csb0    string   `xml:"w:csb0,attr,omitempty"`
	Csb1    string   `xml:"w:csb1,attr,omitempty"`
}

func NewFontTable() *FontTable {
	tb := &FontTable{}
	tb.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	tb.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

	f1 := &font{Name: "Calibri"}
	f1.Panose1 = &fontValue{Val: "020F0502020204030204"}
	f1.Charset = &fontValue{Val: "00"}
	f1.Family = &fontValue{Val: "swiss"}
	f1.Pitch = &fontValue{Val: "variable"}
	f1.Sig = &fontSig{Usb0: "E10002FF", Usb1: "4000ACFF", Usb2: "00000009", Usb3: "00000000", Csb0: "0000019F", Csb1: "00000000"}
	tb.Content = append(tb.Content, f1)

	f2 := &font{Name: "宋体"}
	f2.AltName = &fontValue{Val: "SimSun"}
	f2.Panose1 = &fontValue{Val: "02010600030101010101"}
	f2.Charset = &fontValue{Val: "00"}
	f2.Family = &fontValue{Val: "swiss"}
	f2.Pitch = &fontValue{Val: "variable"}
	f2.Sig = &fontSig{Usb0: "E10002FF", Usb1: "4000ACFF", Usb2: "00000009", Usb3: "00000000", Csb0: "0000019F", Csb1: "00000000"}
	tb.Content = append(tb.Content, f2)

	f3 := &font{Name: "Times New Roman"}
	f3.Panose1 = &fontValue{Val: "02020603050405020304"}
	f3.Charset = &fontValue{Val: "00"}
	f3.Family = &fontValue{Val: "roman"}
	f3.Pitch = &fontValue{Val: "variable"}
	f3.Sig = &fontSig{Usb0: "20002A87", Usb1: "80000000", Usb2: "00000008", Usb3: "00000000", Csb0: "000001FF", Csb1: "00000000"}
	tb.Content = append(tb.Content, f3)

	f4 := &font{Name: "Cambria"}
	f4.Panose1 = &fontValue{Val: "02040503050406030204"}
	f4.Charset = &fontValue{Val: "00"}
	f4.Family = &fontValue{Val: "roman"}
	f4.Pitch = &fontValue{Val: "variable"}
	f4.Sig = &fontSig{Usb0: "A00002EF", Usb1: "4000004B", Usb2: "00000000", Usb3: "00000000", Csb0: "0000019F", Csb1: "00000000"}
	tb.Content = append(tb.Content, f4)

	return tb
}

func (n *FontTable) Save(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(n, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "fontTable.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
