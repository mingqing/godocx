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
