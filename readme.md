xlxs-reader
=====
A fast and lightweight .xlxs or .csv file reader library implemented in Go.

Benchmark 
-------------
Compare with 360EntSecGroup-Skylar/excelize

ReportAllocs

pkg: 360EntSecGroup-Skylar/excelize

-4   	       1	82881214900 ns/op	28014865832 B/op 388529945 allocs/op

pkg:xlsx_reader

-4   	       1	24557851000 ns/op	5508685240 B/op	 155412463 allocs/op


example
-------

    package main
	import "fmt"
	import "github.com/fcode/xlxs_reader"
	func main(){
        file:=`./myexcel.xlsx` //excel file path
        sheetName:="Sheet1"
        cols:=[]string{"colName1","colName2","colName3","..."}
        c:=new(int) //total count
        var i int //current index
        err:=ReadExcel(file,sheetName,&cols,c, func(row []string) {
            //
            i++
            fmt.Printf("processing %v/%v %v\n",i,*c,row)
        })
	}
	
See the go test for more "# xlsx-reader" 
