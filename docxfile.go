package godocx

import (
	"archive/zip"
	//"fmt"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"strings"
)

type DocxFile struct {
	Name   string
	Source *zip.ReadCloser
	Files  []*zipfile
	//Target io.Writer
}

func NewDocxFile(name string) (*DocxFile, error) {
	if name == "" {
		return nil, ErrEmptyName
	}

	return &DocxFile{Name: name,
		Files: make([]*zipfile, 0)}, nil
}

func NewDocxFileFromPath(filePath string) (*DocxFile, error) {
	r, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	//defer r.Close()

	return &DocxFile{Name: path.Base(filePath),
		Source: r,
		Files:  make([]*zipfile, 0)}, nil
}

// 合并docx文件至路径
func (d *DocxFile) CombineTo(docxParentDir, saveDir string) error {
	// 确认saveDir目录存在，以及有相应权限
	err := MustExistAndDir(docxParentDir)
	if err != nil {
		return err
	}
	err = MustExistAndDir(saveDir)
	if err != nil {
		return err
	}

	filepath.Walk(docxParentDir, func(filePath string, info os.FileInfo, err error) error {
		formatPath := strings.Replace(filePath, strings.TrimSpace(path.Join(docxParentDir, " ")), "", -1)
		// 过滤目录
		if err := MustExistAndDir(filePath); err != nil {
			if filePath != "./" {
				fileBody, _ := ioutil.ReadFile(filePath)
				zf, _ := newZipfile(formatPath, fileBody)
				d.Files = append(d.Files, zf)
			}
		}
		return nil
	})

	docxFile, err := os.Create(path.Join(saveDir, d.Name))
	if err != nil {
		return err
	}
	defer docxFile.Close()

	docxWrite := zip.NewWriter(docxFile)
	for _, f := range d.Files {
		t, _ := docxWrite.Create(f.name)
		t.Write(f.body)
	}

	if err := docxWrite.Close(); err != nil {
		return err
	}

	return nil
}

// 分解docx文件至路径
func (d *DocxFile) DecomposeTo(saveDir string) error {
	// 确认saveDir目录存在，以及有相应权限
	err := MustExistAndDir(saveDir)
	if err != nil {
		return err
	}

	// https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
	for _, partFile := range d.Source.File {
		// 排除目录
		if partFile.FileInfo().IsDir() {
			continue
		}

		filePath := path.Join(saveDir, partFile.Name)
		dirPath, _ := path.Split(filePath)
		if err := os.MkdirAll(dirPath, os.ModePerm); err != nil {
			return err
		}

		rc, err := partFile.Open()
		if err != nil {
			return err
		}
		content, _ := ioutil.ReadAll(rc)

		f, err := os.Create(filePath)
		if err != nil {
			return err
		}
		defer f.Close()
		f.Write(content)
		//f.Sync()
	}

	return nil
}
