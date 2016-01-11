package godocx

import (
	"os"
)

func MustExistAndDir(dirPath string) error {
	// 确认dirPath目录存在，以及有相应权限
	dirFile, err := os.Open(dirPath)
	if err != nil {
		return err
	}
	dirStat, err := dirFile.Stat()
	if err != nil {
		return err
	}
	if !dirStat.IsDir() {
		return ErrMustDir
	}

	return nil
}

func MustNotExistAndCreate(dirPath string) error {
	// 确认dirPath目录必须不存在
	_, err := os.Open(dirPath)
	if err != nil {
		if os.IsNotExist(err) {
			os.MkdirAll(dirPath, os.ModePerm)
			return nil
		}
	}

	return ErrExistDir
}
