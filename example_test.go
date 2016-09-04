package godocx

import (
	"fmt"
)

func ExampleNewDocxXml_newDocxXml() {
	d, err := godocx.NewDocxXml("./", "test.docx")
	if err != nil {
		fmt.Println(err)
		return
	}

	d.SaveAndClean()
}

func ExampleDocxFile_combineTo() {
	docx, err := godocx.NewDocxFile("demo.docx")
	if err != nil {
		fmt.Println("err:", err)
	}

	err = docx.CombineTo("./src", "./dist")
	if err != nil {
		fmt.Println("err:", err)
	}
}

func ExampleDocxFile_decombineTo() {
	path := "./dist/demo.docx"
	docx, err := godocx.NewDocxFileFromPath(path)
	if err != nil {
		fmt.Println("err:", err)
	}

	err = docx.DecomposeTo("./src/")
	if err != nil {
		fmt.Println("err:", err)
	}
}
