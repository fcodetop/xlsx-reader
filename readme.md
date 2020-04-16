xlxs-reader
=====
A fast and lightweight .xlxs or .csv file reader library implemented in Go.

一个避免OOM kill的excel读取go的实现

升级了2.0 版本，最终实现低内存并且快速读取.xlxs 文件。

但是比较耗CPU - -!

去掉原来的csv文件读取支持

example 示例
-------

    package main
	import "fmt"
	import "github.com/fcode/xlxs_reader"
	func main(){
        file:=`./myexcel.xlsx` //excel file path
        sheetName:="Sheet1" // zero value,will read the first sheet
        r := Reader(file, sheetName,true) //new a xlxs-reader
        cols,err:= r.Open() 
        
        //or defined columns and validate them
        //cols:=[]string{"colName1","colName2","colName3","..."}
        //err:=r.OpenAndValidCols(cols)   
        	
        	defer r.Close() //reader must be close
        	if(err!=nil){
        		return
        	}
        	fmt.Printf("%v\n", cols)
        	// rowCount,err:=r.GetRowCount() //get total rowcount if need        	
        	err = r.FetchRow(func(row []string) error {
        		fmt.Printf("%v\n", row)
        		d:= row[1] //datetime column
        		//convert excel float value to datetime value
                if  v,err:= strconv.ParseFloat(d,64); err == nil {
                	t:= GetExcelTime(v,true)
                	print(t)
                } 
        		return nil  //retrun an err will break fetch
        	})
       
	}
	
See the go test for more "# xlsx-reader" 
