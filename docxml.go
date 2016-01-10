package godocx

import (
	"encoding/xml"
	"fmt"

	"github.com/mingqing/godocx/word"
)

type docXml struct {
}

func NewDocXml() *docXml {
	return &docXml{}
}

func (d *docXml) Save(dirpath string) error {
	c := newContentType()
	return c.Save(dirpath)
}

func (d *docXml) Test() {
	/*
		c := newContentType()
			output, err := xml.MarshalIndent(c, "", "  ")
			if err != nil {
				fmt.Printf("error: %v\n", err)
			}
	*/
	document := word.NewDocument()
	d.Text1(document)
	d.Text2(document)
	d.Text3(document)
	d.Text4(document)
	d.Text5(document)

	docByte, err := xml.MarshalIndent(document, "", "  ")
	if err != nil {
		fmt.Println("err:", err)
	}
	fmt.Println(xml.Header + string(docByte))
}

func (d *docXml) Text1(document *word.Document) {
	paragh := document.AddParagraph()
	ppr := paragh.AddProperties()
	rpr := ppr.AddRunProperties()
	rpr.Bold(true)
	font := rpr.AddFont()
	font.EastAsia = "黑体"
	run := paragh.AddRunContent()
	rpr2 := run.AddRunProperties()
	rpr2.Bold(true)
	font2 := rpr2.AddFont()
	font2.Ascii = "黑体"
	font2.EastAsia = "黑体"
	font2.Hint = "eastAsia"
	run.Text("绝密★启用前")
}
func (d *docXml) Text2(document *word.Document) {
	paragh := document.AddParagraph()
	ppr := paragh.AddProperties()
	ppr.AddAlign("center")
	rpr := ppr.AddRunProperties()
	rpr.Bold(true)
	font := rpr.AddFont()
	font.EastAsia = "黑体"
	rpr.Fontsize("30")
	rpr.ComplexScriptFontsize("30")
	run := paragh.AddRunContent()
	rpr2 := run.AddRunProperties()
	rpr2.Bold(true)
	rpr2.Fontsize("30")
	rpr2.ComplexScriptFontsize("30")
	font2 := rpr2.AddFont()
	font2.EastAsia = "黑体"
	font2.Hint = "eastAsia"
	run.Text("2015-2016学年度???学校11月月考卷")
}
func (d *docXml) Text3(document *word.Document) {
	paragh := document.AddParagraph()
	ppr := paragh.AddProperties()
	ppr.AddAlign("center")
	rpr := ppr.AddRunProperties()
	rpr.Bold(true)
	font := rpr.AddFont()
	font.EastAsia = "黑体"
	font.AsciiTheme = "minorHAnsi"
	font.HansiTheme = "minorHAnsi"
	rpr.Fontsize("36")
	rpr.ComplexScriptFontsize("36")
	run := paragh.AddRunContent()
	rpr2 := run.AddRunProperties()
	rpr2.Bold(true)
	rpr2.Fontsize("36")
	rpr2.ComplexScriptFontsize("36")
	font2 := rpr2.AddFont()
	font2.EastAsia = "黑体"
	font2.Hint = "eastAsia"
	run.Text("试卷副标题")
}
func (d *docXml) Text4(document *word.Document) {
	paragh := document.AddParagraph()
	ppr := paragh.AddProperties()
	ppr.AddAlign("center")
	run := paragh.AddRunContent()
	rpr2 := run.AddRunProperties()
	font2 := rpr2.AddFont()
	font2.Hint = "eastAsia"
	run.Text("考试范围：xxx；考试时间：100分钟；命题人：xxx")
}
func (d *docXml) Text5(document *word.Document) {
	tbl := document.AddTable()
	pr := tbl.AddProps()
	pr.AddStyle("a1")
	pr.AddWidth("0", "auto")
	pr.AddAlign("center")
	pr.AddBorders()
	pr.AddLayout("fixed")
	pr.AddLook("04A0")
	grid := tbl.AddGrid()
	grid.AddWidth("1000")
	grid.AddWidth("1000")
	grid.AddWidth("1000")
	grid.AddWidth("1000")
	grid.AddWidth("1000")
	tr := tbl.AddRow()
	tr.AddProps().Align("center")

	d.Table1(tr, "题号")
	d.Table1(tr, "一")
	d.Table1(tr, "二")
	d.Table1(tr, "三")
	d.Table1(tr, "总分")
}
func (d *docXml) Table1(tr *word.TableRow, text string) {
	tc := tr.AddCell()
	tcpr := tc.AddProps()
	tcpr.Width("1000", "dxa")
	tcpr.Align("center")
	tcP := tc.AddParagraph()
	tcPPr := tcP.AddProperties()
	tcPPr.AddAlign("center")
	tcPrc := tcP.AddRunContent()
	rcpr := tcPrc.AddRunProperties()
	font := rcpr.AddFont()
	font.Hint = "eastAsia"
	tcPrc.Text(text)
}
