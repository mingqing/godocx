package word

import (
	"encoding/xml"
)

type Document struct {
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
	Body     Body
}

func NewDocument() *Document {
	d := &Document{}

	d.Body = Body{Content: make([]interface{}, 0)}
	paragh := Paragraph{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "00F56D61", RsidRDefault: "00C93DDE"}
	paragh.PPr = &ParagraphProperties{RPr: &RunProperties{}}
	paragh.PPr.RPr.B = &Bold{}
	paragh.PPr.RPr.RFonts = &RunFonts{EastAsia: "黑体"}
	paragh.R = &RunContent{RPr: &RunProperties{}}
	paragh.R.RsidRPr = "00F56D61"
	paragh.R.RPr.RFonts = &RunFonts{Ascii: "黑体", EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.T = "绝密★启用前"
	paragh.R.RPr.B = &Bold{}
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = Paragraph{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "00771D19", RsidRDefault: "00C93DDE"}
	paragh.PPr = &ParagraphProperties{RPr: &RunProperties{}}
	paragh.PPr.Jc = &Jc{Val: "center"}
	paragh.PPr.RPr.RFonts = &RunFonts{EastAsia: "黑体"}
	paragh.PPr.RPr.B = &Bold{}
	paragh.PPr.RPr.Sz = &Fontsize{Val: "30"}
	paragh.PPr.RPr.SzCs = &ComplexScriptFontsize{Val: "30"}
	paragh.R = &RunContent{RPr: &RunProperties{}}
	paragh.R.RsidRPr = "00771D19"
	paragh.R.RPr.RFonts = &RunFonts{Ascii: "黑体", EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.RPr.B = &Bold{}
	paragh.R.RPr.Sz = &Fontsize{Val: "30"}
	paragh.R.RPr.SzCs = &ComplexScriptFontsize{Val: "30"}
	paragh.R.T = "2015-2016学年度???学校11月月考卷"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = Paragraph{RsidR: "00C93DDE", RsidP: "00C93DDE", RsidRPr: "006725CC", RsidRDefault: "00C93DDE"}
	paragh.PPr = &ParagraphProperties{RPr: &RunProperties{}}
	paragh.PPr.Jc = &Jc{Val: "center"}
	paragh.PPr.RPr.RFonts = &RunFonts{AsciiTheme: "minorHAnsi", HansiTheme: "minorHAnsi", EastAsia: "黑体"}
	paragh.PPr.RPr.B = &Bold{}
	paragh.PPr.RPr.Sz = &Fontsize{Val: "36"}
	paragh.PPr.RPr.SzCs = &ComplexScriptFontsize{Val: "36"}
	paragh.R = &RunContent{RPr: &RunProperties{}}
	paragh.R.RsidRPr = "00771D19"
	paragh.R.RPr.RFonts = &RunFonts{EastAsia: "黑体", Hint: "eastAsia"}
	paragh.R.RPr.B = &Bold{}
	paragh.R.RPr.Sz = &Fontsize{Val: "36"}
	paragh.R.RPr.SzCs = &ComplexScriptFontsize{Val: "36"}
	paragh.R.T = "试卷副标题"
	d.Body.Content = append(d.Body.Content, paragh)

	paragh = Paragraph{RsidR: "004D42A0", RsidP: "009C38D0", RsidRDefault: "0063100E"}
	paragh.PPr = &ParagraphProperties{}
	paragh.PPr.Jc = &Jc{Val: "center"}
	paragh.R = &RunContent{RPr: &RunProperties{}}
	paragh.R.RPr.RFonts = &RunFonts{Hint: "eastAsia"}
	paragh.R.T = "考试范围：xxx；考试时间：100分钟；命题人：xxx"
	d.Body.Content = append(d.Body.Content, paragh)

	return d
}
