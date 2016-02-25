package word

import (
	"os"
	"path"
)

type pageHeader struct {
}

func NewHeader() *pageHeader {
	return &pageHeader{}
}

func (h *pageHeader) SaveHeader1(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="000232A6" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a3"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>本卷由系统自动生成，请仔细校对后使用，答案仅供参考。</w:t></w:r></w:p></w:hdr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header1.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageHeader) SaveHeader2(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="000232A6" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a3"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>本卷由系统自动生成，请仔细校对后使用，答案仅供参考。</w:t></w:r></w:p></w:hdr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header2.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageHeader) SaveHeader3(dirpath string) error {
	//	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	//<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="007C4B19" w:rsidRDefault="00304298"><w:pPr><w:pStyle w:val="a3"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto" /></w:pBdr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1029" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:416pt;margin-top:-25pt;width:26pt;height:843pt;z-index:251660288;v-text-anchor:middle" ><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;mso-next-textbox:#_x0000_s1029"><w:txbxContent><w:p w:rsidR="00304298" w:rsidRDefault="00304298" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………内…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:442pt;margin-top:-25pt;width:28pt;height:843pt;z-index:251659264;v-text-anchor:middle" fillcolor="#d8d8d8 [2732]"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="007C4B19" w:rsidP="003E559A"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>※※请※※不※※要※※在※※装※※订※※线※※内※※答※※题※※</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1035" style="position:absolute;left:0;text-align:left;margin-left:442pt;margin-top:-25pt;width:28pt;height:45pt;z-index:351661312" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1035" style="position:absolute;left:0;text-align:left;margin-left:442pt;margin-top:773pt;width:28pt;height:45pt;z-index:351661312" fillcolor="gray [1629]"/></w:pict></w:r><w:r w:rsidR="00207F3C"><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:470pt;margin-top:-25pt;width:26pt;height:843pt;z-index:251658240;v-text-anchor:middle" ><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="007C4B19" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………外…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>`

	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se wp14"><w:p w:rsidR="007C4B19" w:rsidRDefault="00BE540F"><w:pPr><w:pStyle w:val="a3"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s2053" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:424.25pt;margin-top:-41.5pt;width:26pt;height:843pt;z-index:251657216;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;mso-next-textbox:#_x0000_s2053"><w:txbxContent><w:p w:rsidR="00304298" w:rsidRDefault="00B268B5" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………内…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s2050" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:450.25pt;margin-top:-41.5pt;width:28pt;height:843pt;z-index:251655168;v-text-anchor:middle" fillcolor="#d8d8d8 [2732]"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00B268B5" w:rsidP="003E559A"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>※※请※※不※※要※※在※※装※※订※※线※※内※※答※※题※※</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s2059" style="position:absolute;left:0;text-align:left;margin-left:450.25pt;margin-top:-41.5pt;width:28pt;height:45pt;z-index:251659264" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s2049" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:478.25pt;margin-top:-41.5pt;width:26pt;height:843pt;z-index:251653120;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00B268B5" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………外………</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header3.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageHeader) SaveHeader4(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se wp14"><w:p w:rsidR="007C4B19" w:rsidRDefault="003F0172"><w:pPr><w:pStyle w:val="a3"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s2062" style="position:absolute;left:0;text-align:left;margin-left:-63.1pt;margin-top:-41.8pt;width:28pt;height:45pt;z-index:251661312" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s2064" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-34.75pt;margin-top:-41.5pt;width:26pt;height:843pt;z-index:251658240;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="00304298" w:rsidRDefault="00B268B5" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………内…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s2063" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-62.75pt;margin-top:-41.5pt;width:28pt;height:843pt;z-index:251656192;v-text-anchor:middle" fillcolor="#d8d8d8 [2732]"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00B268B5" w:rsidP="003E559A"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>学校</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>:___________</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>姓名：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>___________</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>班级：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>___________</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>考号：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>___________</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s2060" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-88.75pt;margin-top:-41.5pt;width:26pt;height:843pt;z-index:251654144;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00B268B5" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………外…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header4.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageHeader) SaveHeader5(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="007C4B19" w:rsidRDefault="005524F0"><w:pPr><w:pStyle w:val="a3"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:pBdr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1029" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:822pt;margin-top:-43pt;width:26pt;height:730pt;z-index:251657216;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;mso-next-textbox:#_x0000_s1029"><w:txbxContent><w:p w:rsidR="00304298" w:rsidRDefault="00CB653D" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………内…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:848pt;margin-top:-43pt;width:68pt;height:730pt;z-index:251655168;v-text-anchor:middle" fillcolor="#d8d8d8 [2732]"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00CB653D" w:rsidP="003E559A"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>※※请※※不※※要※※在※※装※※订※※线※※内※※答※※题※※</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1035" style="position:absolute;left:0;text-align:left;margin-left:848pt;margin-top:-43pt;width:68pt;height:50pt;z-index:251659264" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1041" style="position:absolute;left:0;text-align:left;margin-left:848pt;margin-top:637pt;width:68pt;height:50pt;z-index:251660288" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:916pt;margin-top:-43pt;width:26pt;height:730pt;z-index:251653120;v-text-anchor:middle"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="00CB653D" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………外…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header5.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageHeader) SaveHeader6(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="007C4B19" w:rsidRDefault="00304298"><w:pPr><w:pStyle w:val="a3"/><w:pBdr><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto" /></w:pBdr></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1029" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-46pt;margin-top:-43pt;width:26pt;height:730pt;z-index:251660288;v-text-anchor:middle" ><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;mso-next-textbox:#_x0000_s1029"><w:txbxContent><w:p w:rsidR="00304298" w:rsidRDefault="00304298" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………内…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-114pt;margin-top:-43pt;width:68pt;height:730pt;z-index:251659264;v-text-anchor:middle" fillcolor="#d8d8d8 [2732]"><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="007C4B19" w:rsidP="003E559A"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>学校:___________姓名：___________班级：___________考号：___________</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1035" style="position:absolute;left:0;text-align:left;margin-left:-114pt;margin-top:-43pt;width:68pt;height:50pt;z-index:351661312" fillcolor="gray [1629]"/></w:pict></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:rect id="_x0000_s1035" style="position:absolute;left:0;text-align:left;margin-left:-114pt;margin-top:637pt;width:68pt;height:50pt;z-index:351661312" fillcolor="gray [1629]"/></w:pict></w:r><w:r w:rsidR="00207F3C"><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;left:0;text-align:left;margin-left:-140pt;margin-top:-43pt;width:26pt;height:730pt;z-index:251658240;v-text-anchor:middle" ><v:textbox style="layout-flow:vertical;mso-layout-flow-alt:bottom-to-top"><w:txbxContent><w:p w:rsidR="007C4B19" w:rsidRDefault="007C4B19" w:rsidP="007C4B19"><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>…………○…………外…………○…………装…………○…………订…………○…………线…………○…………</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p></w:hdr>
`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "header6.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}
