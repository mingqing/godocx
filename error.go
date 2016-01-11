package godocx

import (
	"errors"
)

var (
	ErrMustDir   = errors.New("the path must dir")
	ErrEmptyName = errors.New("the name not empty")
	ErrExistDir  = errors.New("the path exist")
)
