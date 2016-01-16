package word

import (
	"encoding/xml"
	"os"
	"path"
)

type Settings struct {
	XMLName                           xml.Name `xml:"w:settings"`
	XmlnsO                            string   `xml:"w:xmln:o,attr,omitempty"`
	XmlnsR                            string   `xml:"w:xmln:r,attr,omitempty"`
	XmlnsM                            string   `xml:"w:xmln:m,attr,omitempty"`
	XmlnsV                            string   `xml:"w:xmln:v,attr,omitempty"`
	XmlnsW10                          string   `xml:"w:xmln:w10,attr,omitempty"`
	XmlnsW                            string   `xml:"w:xmln:w,attr,omitempty"`
	XmlnsSl                           string   `xml:"w:xmln:sl,attr,omitempty"`
	Zoom                              *zoom
	MirrorMargins                     string `xml:"w:mirrorMargins"`
	BordersDoNotSurroundHeader        string `xml:"w:bordersDoNotSurroundHeader"`
	BordersDoNotSurroundFooter        string `xml:"w:bordersDoNotSurroundFooter"`
	ProofState                        *proofState
	DefaultTabStop                    *settingVal `xml:"w:defaultTabStop"`
	EvenAndOddHeaders                 string      `xml:"w:evenAndOddHeaders"`
	DrawingGridVerticalSpacing        *settingVal `xml:"w:drawingGridVerticalSpacing"`
	DisplayHorizontalDrawingGridEvery *settingVal `xml:"w:displayHorizontalDrawingGridEvery"`
	DisplayVerticalDrawingGridEvery   *settingVal `xml:"w:displayVerticalDrawingGridEvery"`
	CharacterSpacingControl           *settingVal `xml:"w:characterSpacingControl"`
	HdrShapeDefaults                  *hdrShapeDefaults
	FootnotePr                        *footnotePr
	EndnotePr                         *endnotePr
	Compat                            *compat
	Rsids                             *rsids
	MathPr                            *mathPr
	ThemeFontLang                     *themeFontLang
	ClrSchemeMapping                  *clrSchemeMapping
	ShapeDefaults                     *shapeDefaults
	DecimalSymbol                     *settingVal `xml:"w:decimalSymbol"`
	ListSeparator                     *settingVal `xml:"w:listSeparator"`
}

type zoom struct {
	XMLName xml.Name `xml:"w:zoom"`
	Percent string   `xml:"w:percent,attr,omitempty"`
}
type proofState struct {
	XMLName  xml.Name `xml:"w:proofState"`
	Spelling string   `xml:"w:spelling,attr,omitempty"`
	Grammar  string   `xml:"w:grammar,attr,omitempty"`
}
type settingVal struct {
	Val string `xml:"w:val,attr"`
}
type hdrShapeDefaults struct {
	XMLName       xml.Name `xml:"w:hdrShapeDefaults"`
	Shapedefaults *shapedefaults
}
type shapedefaults struct {
	XMLName xml.Name `xml:"o:shapedefaults"`
	Vext    string   `xml:"v:ext,attr"`
	Spidmax string   `xml:"spidmax,attr"`
}
type compat struct {
	XMLName                          xml.Name `xml:"w:compat"`
	SpaceForUL                       string   `xml:"w:spaceForUL"`
	BalanceSingleByteDoubleByteWidth string   `xml:"w:balanceSingleByteDoubleByteWidth"`
	DoNotLeaveBackslashAlone         string   `xml:"w:doNotLeaveBackslashAlone"`
	UlTrailSpace                     string   `xml:"w:ulTrailSpace"`
	DoNotExpandShiftReturn           string   `xml:"w:doNotExpandShiftReturn"`
	AdjustLineHeightInTable          string   `xml:"w:adjustLineHeightInTable"`
	UseFELayout                      string   `xml:"w:useFELayout"`
}

type rsids struct {
	XMLName  xml.Name    `xml:"w:rsids"`
	RsidRoot *settingVal `xml:"w:rsidRoot"`
}

type mathPr struct {
	XMLName    xml.Name    `xml:"w:mathPr"`
	MathFont   *settingVal `xml:"w:mathFont"`
	BrkBin     *settingVal `xml:"w:brkBin"`
	BrkBinSub  *settingVal `xml:"w:brkBinSub"`
	SmallFrac  *settingVal `xml:"w:smallFrac"`
	DispDef    string      `xml:"w:dispDef"`
	LMargin    *settingVal `xml:"w:lMargin"`
	RMargin    *settingVal `xml:"w:rMargin"`
	DefJc      *settingVal `xml:"w:defJc"`
	WrapIndent *settingVal `xml:"w:wrapIndent"`
	IntLim     *settingVal `xml:"w:intLim"`
	NaryLim    *settingVal `xml:"w:naryLim"`
}

type themeFontLang struct {
	XMLName  xml.Name `xml:"w:themeFontLang"`
	Val      string   `xml:"w:val,attr"`
	EastAsia string   `xml:"w:eastAsia,attr"`
}

type clrSchemeMapping struct {
	XMLName           xml.Name `xml:"w:clrSchemeMapping"`
	Bg1               string   `xml:"w:bg1,attr,omitempty"`
	T1                string   `xml:"w:t1,attr,omitempty"`
	Bg2               string   `xml:"w:bg2,attr,omitempty"`
	T2                string   `xml:"w:t2,attr,omitempty"`
	Accent1           string   `xml:"w:accent1,attr,omitempty"`
	Accent2           string   `xml:"w:accent2,attr,omitempty"`
	Accent3           string   `xml:"w:accent3,attr,omitempty"`
	Accent4           string   `xml:"w:accent4,attr,omitempty"`
	Accent5           string   `xml:"w:accent5,attr,omitempty"`
	Accent6           string   `xml:"w:accent6,attr,omitempty"`
	Hyperlink         string   `xml:"w:hyperlink,attr,omitempty"`
	FollowedHyperlink string   `xml:"w:followedHyperlink,attr,omitempty"`
}

type shapeDefaults struct {
	XMLName       xml.Name `xml:"w:shapeDefaults"`
	Shapedefaults *shapedefaults
	Shapelayout   *shapelayout
}
type shapelayout struct {
	XMLName xml.Name `xml:"w:shapelayout"`
	Vext    string   `xml:"v:ext,attr"`
	Idmap   *idmap
}
type idmap struct {
	XMLName xml.Name `xml:"o:idmap"`
	Vext    string   `xml:"v:ext,attr"`
	Data    string   `xml:"data,attr"`
}

type footnotePr struct {
	XMLName xml.Name `xml:"w:footnotePr"`
	Content []*footnote
}
type endnotePr struct {
	XMLName xml.Name `xml:"w:endnotePr"`
	Content []*endnote
}

func NewSettings() *Settings {
	sttng := &Settings{}
	sttng.XmlnsO = "urn:schemas-microsoft-com:office:office"
	sttng.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	sttng.XmlnsM = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	sttng.XmlnsV = "urn:schemas-microsoft-com:vml"
	sttng.XmlnsW10 = "urn:schemas-microsoft-com:office:word"
	sttng.XmlnsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	sttng.XmlnsSl = "http://schemas.openxmlformats.org/schemaLibrary/2006/main"

	sttng.Zoom = &zoom{Percent: "100"}
	//sttng.MirrorMargins = ""
	sttng.ProofState = &proofState{Spelling: "clean", Grammar: "clean"}
	sttng.DefaultTabStop = &settingVal{Val: "420"}
	sttng.DrawingGridVerticalSpacing = &settingVal{Val: "156"}
	sttng.DisplayHorizontalDrawingGridEvery = &settingVal{Val: "0"}
	sttng.DisplayVerticalDrawingGridEvery = &settingVal{Val: "2"}
	sttng.CharacterSpacingControl = &settingVal{Val: "compressPunctuation"}
	sttng.HdrShapeDefaults = &hdrShapeDefaults{Shapedefaults: &shapedefaults{Vext: "edit", Spidmax: "11266"}}
	fn1 := &footnote{Id: "-1"}
	fn2 := &footnote{Id: "0"}
	sttng.FootnotePr = &footnotePr{Content: make([]*footnote, 0)}
	sttng.FootnotePr.Content = append(sttng.FootnotePr.Content, fn1)
	sttng.FootnotePr.Content = append(sttng.FootnotePr.Content, fn2)
	en1 := &endnote{Id: "-1"}
	en2 := &endnote{Id: "0"}
	sttng.EndnotePr = &endnotePr{Content: make([]*endnote, 0)}
	sttng.EndnotePr.Content = append(sttng.EndnotePr.Content, en1)
	sttng.EndnotePr.Content = append(sttng.EndnotePr.Content, en2)
	sttng.Compat = &compat{}
	sttng.Rsids = &rsids{}
	mth := &mathPr{MathFont: &settingVal{Val: "Cambria Math"},
		BrkBin:     &settingVal{Val: "before"},
		BrkBinSub:  &settingVal{Val: "--"},
		SmallFrac:  &settingVal{Val: "off"},
		LMargin:    &settingVal{Val: "0"},
		RMargin:    &settingVal{Val: "0"},
		DefJc:      &settingVal{Val: "centerGroup"},
		WrapIndent: &settingVal{Val: "1440"},
		IntLim:     &settingVal{Val: "subSup"},
		NaryLim:    &settingVal{Val: "undOvr"},
	}
	sttng.MathPr = mth
	sttng.ThemeFontLang = &themeFontLang{Val: "en-US", EastAsia: "zh-CN"}
	sttng.ClrSchemeMapping = &clrSchemeMapping{Bg1: "light1",
		T1: "dark1", Bg2: "light2", T2: "dark2",
		Accent1: "accent1", Accent2: "accent2", Accent3: "accent3",
		Accent4: "accent4", Accent5: "accent5", Accent6: "accent6",
		Hyperlink: "hyperlink", FollowedHyperlink: "followedHyperlink"}

	sl := &shapelayout{Vext: "edit", Idmap: &idmap{Vext: "edit", Data: "2"}}
	sttng.ShapeDefaults = &shapeDefaults{Shapedefaults: &shapedefaults{Vext: "edit", Spidmax: "11266"},
		Shapelayout: sl}

	sttng.DecimalSymbol = &settingVal{Val: "."}
	sttng.ListSeparator = &settingVal{Val: ","}

	return sttng
}

func (n *Settings) Save(dirpath string) error {
	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	output, err := xml.MarshalIndent(n, "", "  ")
	if err != nil {
		return err
	}

	f, err := os.Create(path.Join(fpath, "setting.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(xml.Header)
	f.Write(output)

	return nil
}
