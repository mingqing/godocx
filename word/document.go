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
	Jc      *wordJc
	RPr     *wordRPr
}

// Run props
type wordRPr struct {
	XMLName xml.Name `xml:"w:rPr"`
	RFonts  *wordRFonts
	B       *wordB
	I       *wordI
	Sz      *wordSz
	SzCs    *wordSzCs
}

type wordR struct {
	XMLName xml.Name `xml:"w:r"`
	RsidRPr string   `xml:"w:rsidRPr,attr,omitempty"`
	RPr     *wordRPr `xml:"w:rPr"`
	T       string   `xml:"w:t"`
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

type wordB struct {
	XMLName xml.Name `xml:"w:b"`
}
type wordI struct {
	XMLName xml.Name `xml:"w:i"`
}

type wordRFonts struct {
	XMLName       xml.Name `xml:"w:rFonts"`
	Ascii         string   `xml:"w:ascii,attr,omitempty"`
	EastAsia      string   `xml:"w:eastAsia,attr,omitempty"`
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
	PPr          *wordPPr
	R            *wordR
}

func NewDocument() *document {
	d := &document{}

	d.Body = wordBody{Content: make([]interface{}, 0)}
	paragh := wordP{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "00F56D61", RsidRDefault: "00C93DDE"}
	paragh.PPr = &wordPPr{RPr: &wordRPr{}}
	paragh.PPr.RPr.B = &wordB{}
	paragh.PPr.RPr.RFonts = &wordRFonts{EastAsia: "黑体"}
	paragh.R = &wordR{RPr: &wordRPr{}}
	paragh.R.RsidRPr = "00F56D61"
	paragh.R.RPr.RFonts = &wordRFonts{Ascii: "黑体", EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.T = "绝密★启用前"
	paragh.R.RPr.B = &wordB{}
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "00771D19", RsidRDefault: "00C93DDE"}
	paragh.PPr = &wordPPr{RPr: &wordRPr{}}
	paragh.PPr.Jc = &wordJc{Val: "center"}
	paragh.PPr.RPr.RFonts = &wordRFonts{EastAsia: "黑体"}
	paragh.PPr.RPr.B = &wordB{}
	paragh.PPr.RPr.Sz = &wordSz{Val: "30"}
	paragh.PPr.RPr.SzCs = &wordSzCs{Val: "30"}
	paragh.R = &wordR{RPr: &wordRPr{}}
	paragh.R.RsidRPr = "00771D19"
	paragh.R.RPr.RFonts = &wordRFonts{Ascii: "黑体", EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.RPr.B = &wordB{}
	paragh.R.RPr.Sz = &wordSz{Val: "30"}
	paragh.R.RPr.SzCs = &wordSzCs{Val: "30"}
	paragh.R.T = "2015-2016学年度???学校11月月考卷"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "006725CC", RsidRDefault: "00C93DDE"}
	paragh.PPr = &wordPPr{RPr: &wordRPr{}}
	paragh.PPr.Jc = &wordJc{Val: "center"}
	paragh.PPr.RPr.RFonts = &wordRFonts{AsciiTheme: "minorHAnsi", HansiTheme: "minorHAnsi", EastAsia: "黑体"}
	paragh.PPr.RPr.B = &wordB{}
	paragh.PPr.RPr.Sz = &wordSz{Val: "36"}
	paragh.PPr.RPr.SzCs = &wordSzCs{Val: "36"}
	paragh.R = &wordR{RPr: &wordRPr{}}
	paragh.R.RsidRPr = "00771D19"
	paragh.R.RPr.RFonts = &wordRFonts{EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.RPr.B = &wordB{}
	paragh.R.RPr.Sz = &wordSz{Val: "36"}
	paragh.R.RPr.SzCs = &wordSzCs{Val: "36"}
	paragh.R.T = "试卷副标题"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = wordP{RsidR: "004D42A0", RsidP: "009C38D0", RsidRDefault: "0063100E"}
	paragh.PPr = &wordPPr{}
	paragh.PPr.Jc = &wordJc{Val: "center"}
	paragh.R = &wordR{RPr: &wordRPr{}}
	paragh.R.RPr.RFonts = &wordRFonts{Hint: "eastAsia"}
	paragh.R.T = "考试范围：xxx；考试时间：100分钟；命题人：xxx"
	d.Body.Content = append(d.Body.Content, paragh)

	return d
}
