package word

import (
	"os"
	"path"
)

type pageFooter struct {
}

func NewFooter() *pageFooter {
	return &pageFooter{}
}

func (h *pageFooter) SaveFooter1(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="0064153B" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a4"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>试卷第</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> =</w:instrText></w:r><w:fldSimple w:instr=" page "><w:r><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页，总</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:instrText>=</w:instrText></w:r><w:fldSimple w:instr=" sectionpages "><w:r><w:rPr><w:noProof/></w:rPr><w:instrText>2</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>2</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页</w:t></w:r></w:p></w:ftr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "footer1.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageFooter) SaveFooter2(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="0064153B" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a4"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>试卷第</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> =</w:instrText></w:r><w:fldSimple w:instr=" page "><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页，总</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:instrText>=</w:instrText></w:r><w:fldSimple w:instr=" sectionpages "><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页</w:t></w:r></w:p></w:ftr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "footer2.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageFooter) SaveFooter3(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="0064153B" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a4"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>试卷第</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> =</w:instrText></w:r><w:fldSimple w:instr=" page "><w:r><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页，总</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:instrText>=</w:instrText></w:r><w:fldSimple w:instr=" sectionpages "><w:r><w:rPr><w:noProof/></w:rPr><w:instrText>2</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>2</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页</w:t></w:r></w:p></w:ftr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "footer3.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}

func (h *pageFooter) SaveFooter4(dirpath string) error {
	output := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:p w:rsidR="009E611B" w:rsidRPr="0064153B" w:rsidRDefault="00385C41" w:rsidP="0064153B"><w:pPr><w:pStyle w:val="a4"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>试卷第</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> =</w:instrText></w:r><w:fldSimple w:instr=" page "><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页，总</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:instrText>=</w:instrText></w:r><w:fldSimple w:instr=" sectionpages "><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:instrText>1</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="00B732BA"><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r w:rsidR="00082961"><w:fldChar w:fldCharType="end"/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>页</w:t></w:r></w:p></w:ftr>`

	fpath := path.Join(dirpath, "word")
	os.Mkdir(fpath, os.ModePerm)

	f, err := os.Create(path.Join(fpath, "footer4.xml"))
	if err != nil {
		return err
	}
	defer f.Close()

	f.WriteString(output)

	return nil
}
