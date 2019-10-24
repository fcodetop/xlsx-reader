package xlsx_reader

import (
	"fmt"
	"testing"
)

func TestReaderCsv(t *testing.T) {
	file := `./testfiles/test30k.csv`

	cols := []string{"备注说明", "真实姓名", "*手机号码", "会员昵称", "性别", "出生日期"}

	c := new(int)
	var i int
	err := ReadCsv(file, "gbk", &cols, c, func(row []string) {
		i++
		fmt.Printf("读取进度%v/%v %v\n", i, *c, row)
	})
	if err != nil {
		t.Error(err)
	}
}
