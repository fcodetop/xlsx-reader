package xlsx_reader

import (
	"fmt"
	"testing"
)

func TestReaderExcel(t *testing.T) {
	file := `./testfiles/test30k.xlsx`
	sheetName := "线下会员"
	cols := []string{"会员卡号", "真实姓名", "*手机号码", "会员昵称", "性别", "出生日期",
		"生日", "备注说明"}
	c := new(int)
	var i int
	err := ReadExcel(file, sheetName, &cols, c, func(row []string) {
		i++
		fmt.Printf("读取进度%v/%v %v\n", i, *c, row)
	})
	if err != nil {
		t.Error(err)
	}
}
