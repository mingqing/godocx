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
