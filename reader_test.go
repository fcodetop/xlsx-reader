package xlsx_reader

import (
	"fmt"
	"math"
	"runtime"
	"testing"
)

func TestReader_ReadExlsFast(t *testing.T) {

	file := `E:\testfiles\test300k.xlsx`

	r := Reader(file, "", true)
	cols, err := r.Open()
	defer r.Close()
	if err != nil {
		t.Error(err)
	}
	t.Log(cols)
	err = r.FetchRow(func(row []string) error {
		fmt.Printf("%v\n", row)
		return nil
	})
	mem := runtime.MemStats{}
	runtime.ReadMemStats(&mem)
	t.Logf("TotalAlloc=%vMB", mem.TotalAlloc/1024/1024)
	if err != nil {
		t.Error(err)
	}
}

func TestReader_ReadExlsLowM(t *testing.T) {

	file := `E:\testfiles\test300k.xlsx`

	r := newReader(file, "", true, LowMemery)
	r.Open()
	defer r.Close()
	err := r.FetchRow(func(row []string) error {
		fmt.Printf("%v\n", row)
		return nil
	})
	mem := runtime.MemStats{}
	runtime.ReadMemStats(&mem)
	t.Logf("TotalAlloc=%vB", mem.TotalAlloc)
	if err != nil {
		t.Error(err)
	}
}

func TestReader_OpenAndValidCols(t *testing.T) {

	file := `E:\testfiles\test300k.xlsx`

	r := Reader(file, "", true)
	cols := []string{"真实姓名", "*手机号码", "会员昵称", "性别", "出生日期", "备注说明",
		"积分有效期", "婚姻状况"}
	err := r.OpenAndValidCols(cols)
	defer r.Close()
	if err != nil {
		t.Error(err)
	}
	t.Log(cols)
	err = r.FetchRow(func(row []string) error {
		fmt.Printf("%v\n", row)
		return nil
	})
	mem := runtime.MemStats{}
	runtime.ReadMemStats(&mem)
	t.Logf("TotalAlloc=%vMB", mem.TotalAlloc/1024/1024)
	if err != nil {
		t.Error(err)
	}
}

func TestReader_GetRowCount(t *testing.T) {
	file := `E:\testfiles\test300k.xlsx`

	r := newReader(file, "", false, LowMemery)
	r.Open()
	defer r.Close()
	t.Log(r.GetRowCount())
}

func TestReader_decodeString1(t *testing.T) {
	file := `E:\testfiles\test300k.xlsx`

	r := newReader(file, "", false, LowMemery)
	r.Open()
	defer r.Close()
	r.decodeString1()
}
func TestReader_decodeString(t *testing.T) {
	file := `E:\testfiles\test300k.xlsx`

	r := newReader(file, "", false, LowMemery)
	r.Open()
	defer r.Close()
	r.decodeString()
}

func TestReader_stringToi(t *testing.T) {
	s := "BFB"
	index := 0
	for i := len(s) - 1; i >= 0; i-- {
		x := int(s[i]) - 64
		l := len(s) - i - 1
		index += x*int(math.Pow(26, float64(l))) - 1

	}
	index += len(s) - 1
	t.Log(index)
}

func BenchmarkReader_decodeString1(b *testing.B) {
	file := `E:\testfiles\test300k.xlsx`
	r := newReader(file, "", false, LowMemery)
	r.Open()
	defer r.Close()
	b.ResetTimer()
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		r.decodeString1()
	}
}
func BenchmarkReader_decodeString(b *testing.B) {
	file := `E:\testfiles\test300k.xlsx`
	r := newReader(file, "", false, LowMemery)
	r.Open()
	defer r.Close()
	b.ResetTimer()
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		r.decodeString()
	}
}
