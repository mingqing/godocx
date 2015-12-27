package godocx

import (
//"fmt"
)

type zipfile struct {
	name string // 添加的压缩文件名，排除目录
	body []byte // 压缩的文件内容
}

func newZipfile(name string, body []byte) (*zipfile, error) {
	if name == "" {
		return nil, nil
	}

	return &zipfile{name: name, body: body}, nil
}
