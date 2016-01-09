package word

import (
	"encoding/xml"
)

type document struct {
	XMLName  xml.Name `xml:"w:document"`
	XmlnsO   string   `xml:"xmlns:o,attr,omitempty"`
	XmlnsR   string   `xml:"xmlns:r,attr,omitempty"`
	XmlnsM   string   `xml:"xmlns:m,attr,omitempty"`
	XmlnsV   string   `xml:"xmlns:v,attr,omitempty"`
	XmlnsW   string   `xml:"xmlns:w,attr,omitempty"`
	XmlnsVe  string   `xml:"xmlns:ve,attr,omitempty"`
	XmlnsWp  string   `xml:"xmlns:wp,attr,omitempty"`
	XmlnsW10 string   `xml:"xmlns:w10,attr,omitempty"`
	XmlnsWne string   `xml:"xmlns:wne,attr,omitempty"`
	Body     wordBody
}

type wordBody struct {
	XMLName xml.Name `xml:"w:body"`
	Content []interface{}
}

// Paragraph props
type wordPPr struct {
	XMLName xml.Name `xml:"w:pPr"`
	Jc      wordJc
	RPr     wordRPr
}

// Run props
type wordRPr struct {
	XMLName xml.Name `xml:"w:rPr"`
	B       string   `xml:"w:b,omitempty"`
	I       string   `xml:"w:i,omitempty"`
	RFonts  wordRFonts
	Sz      wordSz
	SzCs    wordSzCs
}

type wordR struct {
	XMLName xml.Name `xml:"w:r"`
	T       string   `xml:"w:t"`
	RPr     wordRPr
}

type wordJc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Font Size
type wordSz struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

// Complex Script Font Size
type wordSzCs struct {
	XMLName xml.Name `xml:"w:szCs"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

type wordRFonts struct {
	XMLName       xml.Name `xml:"w:rFonts"`
	Ascii         string   `xml:"w:ascii,omitempty"`
	EastAsia      string   `xml:"w:eastAsia,omitempty"`
	AsciiTheme    string   `xml:"w:asciiTheme,attr,omitempty"`
	EastAsiaTheme string   `xml:"w:eastAsiaTheme,attr,omitempty"`
	HansiTheme    string   `xml:"w:hAnsiTheme,attr,omitempty"`
	Hint          string   `xml:"w:hint,attr,omitempty"`
}

// Paragraph
type wordP struct {
	XMLName      xml.Name `xml:"w:p"`
	RsidR        string   `xml:"w:rsidR,attr,omitempty"`
	RsidP        string   `xml:"w:rsidP,attr,omitempty"`
	RsidRPr      string   `xml:"w:rsidRPr,attr,omitempty"`
	RsidRDefault string   `xml:"w:rsidRDefault,attr,omitempty"`
	PPr          wordPPr
	R            wordR
}

func NewDocument() *document {
	d := &document{}

	d.Body = wordBody{Content: make([]interface{}, 0)}
	//paragh := wordP{RsidR: "00C93DDE", RsidP: "00F56D61", RsidRPr: "00C93DDE", RsidRDefault: "00C93DDE"}
	paragh := wordP{}
	paragh.PPr.RPr.RFonts.EastAsia = "黑体"
	paragh.PPr.RPr.B = " "
	paragh.R.RPr.RFonts.Ascii = "黑体"
	paragh.R.RPr.RFonts.EastAsia = "黑体"
	paragh.R.RPr.RFonts.Hint = "eastAsia"
	paragh.R.T = "绝密★启用前"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{}
	paragh.PPr.Jc.Val = "center"
	paragh.PPr.RPr.RFonts.EastAsia = "黑体"
	paragh.PPr.RPr.B = " "
	paragh.PPr.RPr.Sz.Val = "30"
	paragh.PPr.RPr.SzCs.Val = "30"
	paragh.R.RPr.RFonts.EastAsia = "黑体"
	paragh.R.RPr.RFonts.Hint = "eastAsia"
	paragh.R.RPr.B = " "
	paragh.R.RPr.Sz.Val = "30"
	paragh.R.RPr.SzCs.Val = "30"
	paragh.R.T = "2015-2016学年度???学校11月月考卷"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{}
	paragh.PPr.Jc.Val = "center"
	paragh.PPr.RPr.RFonts.AsciiTheme = "minorHAnsi"
	paragh.PPr.RPr.RFonts.EastAsia = "黑体"
	paragh.PPr.RPr.RFonts.HansiTheme = "minorHAnsi"
	paragh.PPr.RPr.B = " "
	paragh.PPr.RPr.Sz.Val = "36"
	paragh.PPr.RPr.SzCs.Val = "36"
	paragh.R.RPr.RFonts.EastAsia = "黑体"
	paragh.R.RPr.RFonts.Hint = "eastAsia"
	paragh.R.RPr.B = " "
	paragh.R.RPr.Sz.Val = "36"
	paragh.R.RPr.SzCs.Val = "36"
	paragh.R.T = "试卷副标题"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{}
	paragh.PPr.Jc.Val = "center"
	paragh.R.RPr.RFonts.Hint = "eastAsia"
	paragh.R.T = "考试范围：xxx；考试时间：100分钟；命题人：xxx"
	d.Body.Content = append(d.Body.Content, paragh)

	return d
}
