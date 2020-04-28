package xlsx_reader

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"math"
	"regexp"
	"strconv"
	"strings"
)

type Policy int

const (
	LowMemery = Policy(0) //时间复杂度O(n^2) 空间复杂度O(1)
	Fast      = Policy(1) //时间复杂度O（n） 空间复杂度O(n)

)

var (
	tFlag        = []byte("</t>")
	rowFlag      = []byte("</row>")
	colR         = regexp.MustCompile(`\d+$`)
	ErrFileType  = errors.New("File type must be xlsx")
	ErrSheetName = errors.New("Could not find specific sheet")
	ErrCols      = errors.New("First row does not match Cols")
)

type reader struct {
	fileName      string //xlsx 文件路径及名称
	sheetName     string //读取指定的工作表，如果为空则读取第一个
	policy        Policy //读取策略，快速读取还是小内存读取
	firstRowIsCol bool   //首行数据作为列名

	reader      *zip.ReadCloser
	shareString *zip.File
	sheetData   *zip.File

	sheetReader     io.ReadCloser
	sheetXmlDecoder *xml.Decoder
	cols            []string //firstRowIsCol 为true 时 获取到的列表集合
	columnMaps      map[int]int
	maxIndex        int

	stringCache []string //Fast策略string缓存

	//LowMemery策略 的io指针缓存，一般情况下不需要每次都new
	stringReader io.ReadCloser
	bufReader    *bufio.Reader
	prevIndex    int

	rowCount int //总行数
}

//fileName:xlsx 文件路径及名称,sheetName:读取指定的工作表，如果为空则读取第一个,firstRowIsCol:首行为列名
func Reader(fileName, sheetName string, firstRowIsCol bool) *reader {
	return newReader(fileName, sheetName, firstRowIsCol, Fast)
}
func newReader(fileName, sheetName string, firstRowIsCol bool, policy Policy) *reader {
	return &reader{
		fileName:      fileName,
		sheetName:     sheetName,
		policy:        policy,
		rowCount:      -1,
		firstRowIsCol: firstRowIsCol,
	}
}

//打开要读取的工作表，并根据firstRowIsCol 返回列集合
//如果 firstRowIsCol为false,则cols为nil
//todo 处理列名为空的列
func (this *reader) Open() (cols []string, err error) {
	if !strings.HasSuffix(strings.ToLower(this.fileName), ".xlsx") {
		return nil, ErrFileType
	}
	this.reader, err = zip.OpenReader(this.fileName)
	if err != nil {
		return
	}
	var workbook xlsxWorkbook
	var bookrel xlsxWorkbookRels
	files := this.reader.File
	//解析workbook
	//workbook 文件较小所以可以全量解析
	for _, file := range files {
		if file.Name == "xl/workbook.xml" {
			err = decodeZip(file, &workbook)
			if err != nil {
				return
			}
		} else if file.Name == "xl/_rels/workbook.xml.rels" {
			err = decodeZip(file, &bookrel)
			if err != nil {
				return
			}
		}
	}
	sName := this.getSheetPath(workbook, bookrel)
	if sName == "" {
		err = ErrSheetName
		return
	}
	//得到工作表和字符串存储的xml
	for _, file := range files {
		if file.Name == "xl/sharedStrings.xml" {
			this.shareString = file
		} else if file.Name == sName {
			this.sheetData = file
		}
	}
	//先解析出string
	if this.policy == Fast {
		this.decodeString1()
	}
	this.sheetReader, err = this.sheetData.Open()
	if err != nil {
		return
	}
	this.sheetXmlDecoder = xml.NewDecoder(this.sheetReader)

	//读取首行作为列
	if this.firstRowIsCol {
		var valueFlag int
		var isString bool
	loop:
		for {
			t, _ := this.sheetXmlDecoder.Token()
			if t == nil {
				break
			}
			switch token := t.(type) {
			case xml.StartElement:
				switch token.Name.Local {
				case "row":
					cols = []string{}
				case "c":
					isString = false
					for _, v := range token.Attr {
						if v.Name.Local == "t" && v.Value == "s" {
							isString = true
							break
						}
					}
				case "v":
					valueFlag = 1
				}
			case xml.EndElement:
				name := token.Name.Local
				if name == "row" {
					break loop
				}
			case xml.CharData:
				if valueFlag == 1 {
					value := string([]byte(token))
					if isString {
						index, _ := strconv.Atoi(value)
						if this.policy == Fast {
							value = this.stringCache[index]
						} else {
							value = this.findString(index)
						}
					}
					cols = append(cols, value)
				}
			}
		}
		this.cols = cols
		this.columnMaps = make(map[int]int, len(cols))
		for i := 0; i < len(cols); i++ {
			this.columnMaps[i] = i
		}
		this.maxIndex = len(cols) - 1
	}
	return
}

//打开要读取的工作表，并根据输入的cols校验excel模板是否正确
//firstRowIsCol 参数必须为true
func (this *reader) OpenAndValidCols(cols []string) error {
	if !this.firstRowIsCol {
		return errors.New("firstRowIsCol must be true")
	}
	if _, err := this.Open(); err != nil {
		return err
	}
	if err := this.checkCols(cols); err != nil {
		return err
	}
	this.cols = cols
	return nil
}

func (this *reader) checkCols(cols []string) error {
	l := len(cols)
	if l > len(this.cols) {
		return ErrCols
	}
	this.columnMaps = make(map[int]int, l)
	var isFound bool
	for i, v := range cols {
		isFound = false
		for j, c := range this.cols {
			if v == c {
				this.columnMaps[j] = i
				if this.maxIndex < j {
					this.maxIndex = j
				}
				isFound = true
				break
			}
		}
		//第一行中不包含传入的列
		if !isFound {
			return ErrCols
		}
	}
	return nil
}

func (this *reader) Close() error {
	if this.stringReader != nil {
		this.stringReader.Close()
	}
	if this.sheetReader != nil {
		this.sheetReader.Close()
	}
	if this.reader != nil {
		return this.reader.Close()
	}
	return nil
}

//逐行读取，如果rowAction中返回 err!=nil 则中断
func (this *reader) FetchRow(rowAction func(row []string) error) (err error) {
	//解析工作表，这里如果全量解析内部使用递归算法，所以只能逐行解析，避免OOM kill
	//flag: 0 ignore,1 elementStart, 2 elementEnd
	var rowFlag, valueFlag int8
	var prevColIndex, colIndex, realIndex int
	var isString, ok bool
	var row []string
	//逐行读取sheet
loop:
	for {
		t, _ := this.sheetXmlDecoder.Token()
		if t == nil {
			break
		}
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "row":
				rowFlag = 1
				prevColIndex = 0
				if this.firstRowIsCol {
					row = make([]string, len(this.cols))

				} else {
					row = []string{}
				}
			case "c":
				if rowFlag == 1 {
					isString = false
					for _, v := range token.Attr {
						if v.Name.Local == "t" && v.Value == "s" {
							isString = true
						} else if v.Name.Local == "r" {
							colIndex = getIndex(v.Value)
							//忽略超过指定列的数据
							if this.firstRowIsCol && colIndex > this.maxIndex {
								rowFlag = 0
							}
						}
					}
				}
			case "v":
				if rowFlag == 1 {
					valueFlag = 1
				}
			}
		case xml.EndElement:
			switch token.Name.Local {
			case "row":
				rowFlag = 2
			case "c":
				prevColIndex = colIndex
			case "v":
				valueFlag = 2
			case "sheetData":
				break loop
			}
		case xml.CharData:
			if valueFlag == 1 {
				if this.firstRowIsCol {
					if realIndex, ok = this.columnMaps[colIndex]; !ok {
						break
					}
				}
				value := string([]byte(token))
				if isString {
					i, _ := strconv.Atoi(value)
					if this.policy == Fast {
						value = this.stringCache[i]
					} else {
						value = this.findString(i)
					}
				}
				if this.firstRowIsCol {
					row[realIndex] = value
				} else {
					for i := prevColIndex; i <= colIndex; i++ {
						if i == colIndex {
							row = append(row, value)
						} else {
							row = append(row, "")
						}
					}
				}
			}
		}

		if rowFlag == 2 {
			if er := rowAction(row); er != nil {
				return er
			}
		}
	}

	return
}

func getIndex(colId string) int {
	s := colR.ReplaceAllString(colId, "")
	index := 0
	for i := len(s) - 1; i >= 0; i-- {
		x := int(s[i]) - 64
		l := len(s) - i - 1
		index += x*int(math.Pow(26, float64(l))) - 1
	}
	index += len(s) - 1
	return index
}

//获取总行数,如果需要时则获取
func (this *reader) GetRowCount() (c int, err error) {
	if this.rowCount > -1 {
		c = this.rowCount
		return
	}
	r, err := this.sheetData.Open()
	if err != nil {
		return
	}
	defer r.Close()
	buf := bufio.NewReader(r)
	for {
		bs, err := buf.ReadBytes('>')
		if err != nil {
			break
		}
		if bytes.Contains(bs, rowFlag) {
			c++
		}
	}
	this.rowCount = c
	return
}

func (this *reader) getSheetPath(workbook xlsxWorkbook, bookrel xlsxWorkbookRels) string {
	var rid string
	if this.sheetName != "" {
		for _, sheet := range workbook.Sheets.Sheet {
			if sheet.Name == this.sheetName {
				rid = sheet.ID
				break
			}
		}
	}
	if rid == "" {
		rid = workbook.Sheets.Sheet[0].ID
	}
	for _, rel := range bookrel.Relationships {
		if rel.ID == rid {
			return "xl/" + rel.Target
		}
	}
	return ""
}

//全量解析XML
func decodeZip(f *zip.File, v interface{}) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	decoder := xml.NewDecoder(rc)
	err = decoder.Decode(v)
	return err
}

//时间复杂度最大，空间复杂度最小的 字符串查找，每次遍历不缓存
//todo 如果文件较大可以使用分片多协程查找
func (this *reader) findString(i int) string {
	//保存上一次的查找位置，一般情况下，不需要重头开始查
	if this.bufReader == nil || i <= this.prevIndex {
		if this.stringReader != nil {
			this.stringReader.Close()
		}
		rc, err := this.shareString.Open()
		if err != nil {
			return ""
		}
		this.stringReader = rc

		this.bufReader = bufio.NewReader(this.stringReader)
		this.prevIndex = 0
	}
	for {
		bs, err := this.bufReader.ReadBytes('>')
		if err != nil {
			break
		}
		if in := bytes.Index(bs, tFlag); in > -1 {
			if i == this.prevIndex {
				this.prevIndex++
				return string(bs[0:in])
			}
			this.prevIndex++
		}
	}
	return ""
}

//解析shareString到缓存
func (this *reader) decodeString() error {
	rc, err := this.shareString.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	index := 0
	buf := bufio.NewReader(rc)
	isMake := false
	reg := regexp.MustCompile(`uniqueCount="(\d+)"`)
	for {
		bs, err := buf.ReadBytes('>')
		if err != nil {
			break
		}
		if !isMake {
			r := reg.FindSubmatch(bs)
			if len(r) > 1 {
				c, err := strconv.Atoi(string(r[1]))
				if err != nil {
					return err
				}
				this.stringCache = make([]string, c)
				isMake = true
			}
		}
		if i := bytes.Index(bs, tFlag); i > -1 {
			this.stringCache[index] = string(bs[0:i])
			index++
		}
	}
	return nil
}

//解析shareString到缓存（标准库xml解析）
func (this *reader) decodeString1() error {
	rc, err := this.shareString.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	var valueFlag int
	index := 0

loop:
	for {
		t, _ := d.Token()
		if t == nil {
			break
		}
		switch token := t.(type) {
		case xml.StartElement:
			name := token.Name.Local
			if name == "sst" {
				for _, v := range token.Attr {
					if v.Name.Local == "uniqueCount" {
						c, _ := strconv.Atoi(v.Value)
						this.stringCache = make([]string, c)
						break
					}
				}
			} else if name == "t" {
				valueFlag = 1
			}
		case xml.EndElement:
			name := token.Name.Local
			if name == "t" {
				valueFlag = 2
				index++
			} else if name == "sst" {
				break loop
			}
		case xml.CharData:
			if valueFlag == 1 {
				this.stringCache[index] = string([]byte(token))
			}
		}
	}
	return nil
}
