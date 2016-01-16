package godocx

import (
	"encoding/xml"
	"fmt"
	"path/filepath"

	"github.com/mingqing/godocx/word"
)

type docXml struct {
	Dir  string
	Name string
}

func NewDocXml(fpath string) (*docXml, error) {
	d := &docXml{Dir: filepath.Dir(fpath), Name: filepath.Base(fpath)}
	if err := MustNotExistAndCreate(d.Dir); err != nil {
		fmt.Println("err:", err)
		return nil, err
	}

	fmt.Printf("dir: {%s} name: {%s}\n", d.Dir, d.Name)

	d.ContentType(d.Dir)
	d.DocProps(d.Dir)
	d.Rels(d.Dir)
	d.FootEndNotes(d.Dir)
	if err := d.FontTable(d.Dir); err != nil {
		fmt.Println("font table:", err)
	}
	if err := d.Settings(d.Dir); err != nil {
		fmt.Println("settings:", err)
	}
	if err := d.Styles(d.Dir); err != nil {
		fmt.Println("styles:", err)
	}
	return d, nil
}

func (d *docXml) Document() *word.Document {
	return word.NewDocument()
}
func (d *docXml) ContentType(dirpath string) error {
	c := newContentType()
	return c.Save(dirpath)
}
func (d *docXml) DocProps(dirpath string) error {
	app := newDocPropsApp()
	app.Save(dirpath)

	core := newDocPropsCore()
	core.Save(dirpath)

	return nil
}
func (d *docXml) Rels(dirpath string) error {
	rels := newRelationships()
	rels.Save(dirpath)
	return nil
}
func (d *docXml) FootEndNotes(dirpath string) error {
	f := word.NewFootnotes()
	f.Save(dirpath)

	e := word.NewEndnotes()
	e.Save(dirpath)
	return nil
}
func (d *docXml) FontTable(dirpath string) error {
	f := word.NewFontTable()
	return f.Save(dirpath)
}
func (d *docXml) Settings(dirpath string) error {
	f := word.NewSettings()
	if err := f.Save(dirpath); err != nil {
		return err
	}

	if err := word.NewWebSettings().Save(dirpath); err != nil {
		return err
	}

	return nil
}
func (d *docXml) Styles(dirpath string) error {
	f := word.NewStyles()
	return f.Save(dirpath)
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
	//d.Text2(document)
	//d.Text3(document)
	//d.Text4(document)
	//d.Text5(document)
	//d.Text6(document)
	//d.Text7(document)
	//d.Text8(document)

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
	border := pr.AddBorders()
	temp := &word.Border{Val: "single", Sz: "4", Space: "0", Color: "000000", ThemeColor: "text1"}
	border.Top = temp
	border.Left = temp
	border.Bottom = temp
	border.Right = temp
	border.InsideH = temp
	border.InsideV = temp
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

	tr2 := tbl.AddRow()
	tr2pr := tr2.AddProps()
	tr2pr.Align("center")
	tr2pr.Height("520")
	d.Table1(tr2, "得分")
	d.Table1(tr2, "")
	d.Table1(tr2, "")
	d.Table1(tr2, "")
	d.Table1(tr2, "")
}
func (d *docXml) Table1(tr *word.TableRow, text string) {
	tc := tr.AddCell()
	tcpr := tc.AddProps()
	tcpr.Width("1000", "dxa")
	tcpr.Align("center")
	tcP := tc.AddParagraph()
	tcPPr := tcP.AddProperties()
	tcPPr.AddAlign("center")
	if text != "" {
		tcPrc := tcP.AddRunContent()
		rcpr := tcPrc.AddRunProperties()
		font := rcpr.AddFont()
		font.Hint = "eastAsia"
		tcPrc.Text(text)
	}
}
func (d *docXml) Text6(document *word.Document) {
	d.defaultText(document, "注意事项：")
	d.defaultText(document, "1．答题前填写好自己的姓名、班级、考号等信息")
	d.defaultText(document, "2．请将答案正确填写在答题卡上")
	d.defaultText(document, "")
}
func (d *docXml) defaultText(document *word.Document, text string) {
	paragh := document.AddParagraph()
	run := paragh.AddRunContent()
	rpr := run.AddRunProperties()
	font := rpr.AddFont()
	font.Hint = "eastAsia"
	run.Text(text)
}
func (d *docXml) Text7(document *word.Document) {
	paragh := document.AddParagraph()
	ppr := paragh.AddProperties()
	ppr.AddAlign("center")
	rpr := ppr.AddRunProperties()
	rpr.Bold(true)
	font := rpr.AddFont()
	font.AsciiTheme = "majorEastAsia"
	font.EastAsiaTheme = "majorEastAsia"
	font.HansiTheme = "majorEastAsia"
	rpr.Fontsize("24")
	rpr.ComplexScriptFontsize("24")
	run := paragh.AddRunContent()
	rpr2 := run.AddRunProperties()
	rpr2.Bold(true)
	rpr2.Fontsize("24")
	rpr2.ComplexScriptFontsize("24")
	font2 := rpr2.AddFont()
	font2.AsciiTheme = "majorEastAsia"
	font2.EastAsiaTheme = "majorEastAsia"
	font2.HansiTheme = "majorEastAsia"
	font2.Hint = "eastAsia"
	run.Text("第I卷（选择题）")

	d.defaultText(document, "")
}
func (d *docXml) Text8(document *word.Document) {
	tbl := document.AddTable()
	pr := tbl.AddProps()
	pr.AddStyle("a7")
	pr.AddWidth("0", "auto")
	border := pr.AddBorders()
	temp := &word.Border{Val: "none", Sz: "0", Space: "0", Color: "auto"}
	border.Top = temp
	border.Left = temp
	border.Bottom = temp
	border.Right = temp
	border.InsideH = temp
	border.InsideV = temp
	pr.AddLook("04A0")
	grid := tbl.AddGrid()
	grid.AddWidth("2184")
	grid.AddWidth("3200")
	tr := tbl.AddRow()
	tc := tr.AddCell()
	tcpr := tc.AddProps()
	tcpr.Width("0", "auto")

	tbl2 := word.NewTable()
	pr2 := tbl2.AddProps()
	pr2.AddStyle("a7")
	pr2.AddWidth("1938", "dxa")
	border2 := pr2.AddBorders()
	temp2 := &word.Border{Val: "signle", Sz: "12", Space: "0", Color: "auto"}
	border2.Top = temp2
	border2.Left = temp2
	border2.Bottom = temp2
	border2.Right = temp2
	border2.InsideH = temp2
	border2.InsideV = temp2
	pr2.AddLook("04A0")
	grid2 := tbl2.AddGrid()
	grid2.AddWidth("969")
	grid2.AddWidth("969")
	tr2 := tbl2.AddRow()
	tr2pr := tr2.AddProps()
	tr2pr.Height("274")
	d.Table2(tr2, "评卷人")
	d.Table2(tr2, "得分")
	tr3 := tbl2.AddRow()
	tr3pr := tr3.AddProps()
	tr3pr.Height("549")
	d.Table2(tr3, "")
	d.Table2(tr3, "")

	tc.Content = make([]interface{}, 0)
	tc.Content = append(tc.Content, tbl2)
	p := &word.Paragraph{}
	tc.Content = append(tc.Content, p)

	tc2 := tr.AddCell()
	tcpr2 := tc2.AddProps()
	tcpr2.Width("0", "auto")
	tcpr2.Align("center")
	tcP := tc2.AddParagraph()
	tcPPr := tcP.AddProperties()
	//tcPPr.AddAlign("center")
	tcPPr.AddRunProperties().Bold(true)
	tcPrc := tcP.AddRunContent()
	rcpr := tcPrc.AddRunProperties()
	font := rcpr.AddFont()
	font.Hint = "eastAsia"
	rcpr.Bold(true)
	tcPrc.Text("一、选择题（题型注释）")
}
func (d *docXml) Table2(tr *word.TableRow, text string) {
	tc := tr.AddCell()
	tcpr := tc.AddProps()
	tcpr.Width("969", "dxa")
	tcP := tc.AddParagraph()
	tcPPr := tcP.AddProperties()
	tcPPr.AddAlign("center")
	if text != "" {
		tcPrc := tcP.AddRunContent()
		rcpr := tcPrc.AddRunProperties()
		font := rcpr.AddFont()
		font.Hint = "eastAsia"
		tcPrc.Text(text)
	}
}
