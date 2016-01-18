package word

import (
	"encoding/xml"
	"os"
	"path"
	"strconv"
)

type Styles struct {
	XMLName      xml.Name `xml:"w:styles"`
	XmlnsR       string   `xml:"xmln:r,attr,omitempty"`
	XmlnsW       string   `xml:"xmln:w,attr,omitempty"`
	DocDefaults  *docDefaults
	LatentStyles *latentStyles
	Style        []*wStyle
}

type docDefaults struct {
	XMLName    xml.Name `xml:"w:docDefaults"`
	RPrDefault *rPrDefault
	PPrDefault string `xml:"w:pPrDefault"`
}
type rPrDefault struct {
	XMLName xml.Name `xml:"w:rPrDefault"`
	RPr     *RunProperties
}

type latentStyles struct {
	XMLName           xml.Name `xml:"w:latentStyles"`
	DefLockedState    string   `xml:"w:defLockedState,attr,omitempty"`
	DefUIPriority     string   `xml:"w:defUIPriority,attr,omitempty"`
	DefSemiHidden     string   `xml:"w:defSemiHidden,attr,omitempty"`
	DefUnhideWhenUsed string   `xml:"w:defUnhideWhenUsed,attr,omitempty"`
	DefQFormat        string   `xml:"w:defQFormat,attr,omitempty"`
	Count             string   `xml:"w:count,attr,omitempty"`
	LsdException      []*lsdException
}
type lsdException struct {
	XMLName        xml.Name `xml:"w:lsdException"`
	Name           string   `xml:"w:name,attr,omitempty"`
	SemiHidden     string   `xml:"w:semiHidden,attr,omitempty"`
	UiPriority     string   `xml:"w:uiPriority,attr,omitempty"`
	UnhideWhenUsed string   `xml:"w:unhideWhenUsed,attr,omitempty"`
	QFormat        string   `xml:"w:qFormat,attr,omitempty"`
}

func newLsdException(name, sh, up, uh, qf string) *lsdException {
	return &lsdException{Name: name, SemiHidden: sh, UiPriority: up, UnhideWhenUsed: uh, QFormat: qf}
}

type wStyle struct {
	XMLName        xml.Name    `xml:"w:style"`
	Type           string      `xml:"w:type,attr"`
	Default        string      `xml:"w:default,attr"`
	StyleId        string      `xml:"w:styleId,attr"`
	Name           *settingVal `xml:"w:name"`
	BasedOn        *settingVal `xml:"w:basedOn"`
	Link           *settingVal `xml:"w:link"`
	UiPriority     *settingVal `xml:"w:uiPriority"`
	SemiHidden     string      `xml:"w:semiHidden,omitempty"`
	UnhideWhenUsed string      `xml:"w:unhideWhenUsed,omitempty"`
	QFormat        string      `xml:"w:qFormat,omitempty"`
	Rsid           *settingVal `xml:"w:rsid"`
	PPr            *ParagraphProperties
	RPr            *RunProperties
}

func newStyle(typ, dft, sid, name string) *wStyle {
	return &wStyle{Type: typ, Default: dft, StyleId: sid, Name: &settingVal{Val: name}}
}

func NewStyles() *Styles {
	styles := &Styles{}
	/*
		styles.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		styles.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

		rpr := NewRunProperties()
		rt := rpr.AddFont()
		rt.AsciiTheme = "minorHAnsi"
		rt.EastAsiaTheme = "minorEastAsia"
		rt.HansiTheme = "minorHAnsi"
		rt.Cstheme = "minorBidi"

		lg := rpr.AddLang()
		lg.Val = "en-US"
		lg.EastAsia = "zh-CN"
		lg.Bidi = "ar-SA"

		rpr.Kern("2")
		rpr.Fontsize("21")
		rpr.ComplexScriptFontsize("22")

		rPrD := &rPrDefault{RPr: rpr}
		styles.DocDefaults = &docDefaults{RPrDefault: rPrD}

		// Step 2
		ls := &latentStyles{DefLockedState: "0",
			DefUIPriority:     "99",
			DefSemiHidden:     "1",
			DefUnhideWhenUsed: "1",
			DefQFormat:        "0",
			Count:             "267",
			LsdException:      make([]*lsdException, 0)}

		ls.addCommon()
		ls.addHeadAndToc()
		ls.addOther()
		styles.LatentStyles = ls

		// Step 3
		styles.addStyle()
	*/

	return styles
}

func (ls *latentStyles) addCommon() {
	ls.LsdException = append(ls.LsdException, newLsdException("Placeholder Text", "", "", "0", ""))
	ls.LsdException = append(ls.LsdException, newLsdException("Revision", "", "", "0", ""))
	ls.LsdException = append(ls.LsdException, newLsdException("Default Paragraph Font", "", "1", "", ""))
	ls.LsdException = append(ls.LsdException, newLsdException("caption", "", "35", "", "1"))
	ls.LsdException = append(ls.LsdException, newLsdException("Bibliography", "", "37", "", ""))
	ls.LsdException = append(ls.LsdException, newLsdException("Table Grid", "0", "59", "0", ""))

	lsNames := []string{"Title", "Subtitle"}
	up := 10
	for _, name := range lsNames {
		ls.LsdException = append(ls.LsdException, newLsdException(name, "0", strconv.Itoa(up), "0", "1"))
		up += 1
	}

	lsNames = []string{"Subtle Emphasis", "Emphasis", "Intense Emphasis", "Strong"}
	up = 19
	for _, name := range lsNames {
		ls.LsdException = append(ls.LsdException, newLsdException(name, "0", strconv.Itoa(up), "0", "1"))
		up += 1
	}

	lsNames = []string{"Quote", "Intense Quote", "Subtle Reference", "Intense Reference", "Book Title", "List Paragraph"}
	up = 29
	for _, name := range lsNames {
		ls.LsdException = append(ls.LsdException, newLsdException(name, "0", strconv.Itoa(up), "0", "1"))
		up += 1
	}
}
func (ls *latentStyles) addHeadAndToc() {
	for i := 1; i < 10; i++ {
		ls.LsdException = append(ls.LsdException, newLsdException("heading "+strconv.Itoa(i), "", "9", "", "1"))
		ls.LsdException = append(ls.LsdException, newLsdException("toc "+strconv.Itoa(i), "", "39", "", "1"))
	}
	ls.LsdException = append(ls.LsdException, newLsdException("TOC Heading", "", "39", "", "1"))
}
func (ls *latentStyles) addOther() {
	for index := 0; index < 7; index++ {
		if index == 0 {
			ls.LsdException = append(ls.LsdException, newLsdException("Light Shading", "", "60", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Light List", "", "61", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Light Grid", "", "62", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Shading 1", "", "63", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Shading 2", "", "64", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium List 1", "", "65", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium List 2", "", "66", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 1", "", "67", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 2", "", "68", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 3", "", "69", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Dark List", "", "70", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful Shading", "", "71", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful List", "", "72", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful Grid", "", "73", "", ""))
		} else {
			ls.LsdException = append(ls.LsdException, newLsdException("Light Shading Accent ", "", "60", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Light List Accent ", "", "61", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Light Grid Accent ", "", "62", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Shading 1 Accent ", "", "63", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Shading 2 Accent ", "", "64", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium List 1 Accent ", "", "65", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium List 2 Accent ", "", "66", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 1 Accent ", "", "67", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 2 Accent ", "", "68", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Medium Grid 3 Accent ", "", "69", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Dark List Accent ", "", "70", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful Shading Accent ", "", "71", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful List Accent ", "", "72", "", ""))
			ls.LsdException = append(ls.LsdException, newLsdException("Colorful Grid Accent ", "", "73", "", ""))
		}
	}
}

func (n *Styles) addStyle() {
	s1 := newStyle("paragraph", "1", "a", "Normal")
	s1.Rsid = &settingVal{Val: "002A2386"}
	s1.PPr = NewParagraphProperties()
	s1.PPr.AddAlign("both")

	s2 := newStyle("character", "1", "a0", "Default Paragraph Font")
	s2.UiPriority = &settingVal{Val: "1"}

	s3 := newStyle("table", "1", "a1", "Normal Table")
	s3.UiPriority = &settingVal{Val: "99"}
	//s3.tablPr

	s4 := newStyle("numbering", "1", "a2", "No List")
	s4.UiPriority = &settingVal{Val: "99"}

	s5 := newStyle("paragraph", "1", "a3", "header")
	s5.BasedOn = &settingVal{Val: "a"}
	s5.Link = &settingVal{Val: "Char"}
	s5.UiPriority = &settingVal{Val: "99"}
	s5.Rsid = &settingVal{Val: "00AD3992"}
	s5.PPr = NewParagraphProperties()
	s5.PPr.AddAlign("center")
	//s5.RPr

	n.Style = append(n.Style, s1)
	n.Style = append(n.Style, s2)
	n.Style = append(n.Style, s3)
	n.Style = append(n.Style, s4)
	n.Style = append(n.Style, s5)
}

func (n *Styles) Save2(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(n, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "styles.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}

func (n *Styles) Save(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="002A2386"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="a2"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="paragraph" w:styleId="a3"><w:name w:val="header"/><w:basedOn w:val="a"/><w:link w:val="Char"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rsid w:val="00AD3992"/><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char"><w:name w:val="页眉 Char"/><w:basedOn w:val="a0"/><w:link w:val="a3"/><w:uiPriority w:val="99"/><w:rsid w:val="00AD3992"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a4"><w:name w:val="footer"/><w:basedOn w:val="a"/><w:link w:val="Char0"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00AD3992"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char0"><w:name w:val="页脚 Char"/><w:basedOn w:val="a0"/><w:link w:val="a4"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00AD3992"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a5"><w:name w:val="Balloon Text"/><w:basedOn w:val="a"/><w:link w:val="Char1"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00AD3992"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char1"><w:name w:val="批注框文本 Char"/><w:basedOn w:val="a0"/><w:link w:val="a5"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00AD3992"/><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="a6"><w:name w:val="No Spacing"/><w:link w:val="Char2"/><w:uiPriority w:val="1"/><w:qFormat/><w:rsid w:val="00AD3992"/><w:rPr><w:kern w:val="0"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="Char2"><w:name w:val="无间隔 Char"/><w:basedOn w:val="a0"/><w:link w:val="a6"/><w:uiPriority w:val="1"/><w:rsid w:val="00AD3992"/><w:rPr><w:kern w:val="0"/><w:sz w:val="22"/></w:rPr></w:style></w:styles>
    `
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "styles.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}
