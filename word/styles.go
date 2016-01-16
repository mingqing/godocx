package word

import (
	"encoding/xml"
	"os"
	"path"
	"strconv"
)

type Styles struct {
	XMLName      xml.Name `xml:"w:styles"`
	XmlnsR       string   `xml:"w:xmln:r,attr,omitempty"`
	XmlnsW       string   `xml:"w:xmln:w,attr,omitempty"`
	DocDefaults  *docDefaults
	LatentStyles *latentStyles
	//Style        *style
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

func NewStyles() *Styles {
	styles := &Styles{}
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

func (n *Styles) Save(dirpath string) error {
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
