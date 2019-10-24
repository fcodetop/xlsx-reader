package xlsx_reader

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"math"
	"regexp"
	"strconv"
	"time"
)

const matchColsErr = "sheetData did not match cols."

//if sheetName is zero string will retrun the first sheet
func ReadExcel(filePath, sheetName string, cols *[]string, totalCount *int, rowAction func(row []string)) error {
	sheet, shareStr, err := deCompress(filePath, sheetName)
	if err != nil {
		return err
	}

	sheetData := sheet.SheetData
	r := len(sheetData.Row) - 1
	*totalCount = r
	if r > 0 {
		header := &sheetData.Row[0].C
		fillStr(header, shareStr)
		m, err := checkCol(cols, header)
		if err != nil {
			return err
		}
		var unMatch bool
		var xr []xlsxC
		c := len(*cols)
		row := make([]string, c)
		for i := 1; i <= r; i++ {
			xr = sheetData.Row[i].C
			for j := 0; j < c; j++ {
				unMatch = true
				for k := 0; k < len(xr); k++ {
					v := xr[k]
					if v.R == fmt.Sprint(m[j], i+1) {
						if v.T == "s" {
							index, _ := strconv.Atoi(v.V)
							row[j] = shareStr.SI[index].T
						} else {
							row[j] = v.V
						}
						unMatch = false
						break
					}
				}
				if unMatch {
					row[j] = ""
				}
			}
			rowAction(row)
		}
	}
	return nil
}

//file the string cell value
func fillStr(cells *[]xlsxC, shareStr *xlsxSST) {
	for i, v := range *cells {
		if v.T == "s" {
			index, _ := strconv.Atoi(v.V)
			(*cells)[i].V = shareStr.SI[index].T
		}
	}
}

//check cols if match
func checkCol(cols *[]string, header *[]xlsxC) ([]string, error) {
	l := len(*cols)
	if l > len(*header) {
		return nil, errors.New(matchColsErr)
	}
	r := regexp.MustCompile(`\d+$`)
	m := make([]string, l)
	for i, v := range *cols {
		for _, c := range *header {
			if v == c.V {
				m[i] = r.ReplaceAllString(c.R, "")
				break
			}
		}

		if m[i] == "" {
			return nil, errors.New(matchColsErr)
		}
	}
	return m, nil
}

func deCompress(file, sheetName string) (sheet *xlsxWorksheet, sst *xlsxSST, err error) {
	reader, err := zip.OpenReader(file)
	if err != nil {
		return
	}
	defer reader.Close()
	var workbook xlsxWorkbook
	var bookrel xlsxWorkbookRels
	files := reader.File
	//resolve workbook
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
	sName := getSheetPath(&workbook, &bookrel, sheetName)
	for _, file := range files {
		if file.Name == "xl/sharedStrings.xml" {
			sst = new(xlsxSST)
			err = decodeZip(file, sst)
			if err != nil {
				return
			}
		} else if file.Name == sName {
			sheet = new(xlsxWorksheet)
			err = decodeZip(file, sheet)
			if err != nil {
				return
			}
		}
	}
	return
}

// deccode zip xml file content as struct.
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

//get sheet xml file name
//if sheetName is zero string will retrun the first sheet
func getSheetPath(workbook *xlsxWorkbook, bookrel *xlsxWorkbookRels, sheetName string) string {
	var rid string
	if sheetName != "" {
		for _, sheet := range workbook.Sheets.Sheet {
			if sheet.Name == sheetName {
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

// julianDateToGregorianTime provides a function to convert julian date to
// gregorian time.
func julianDateToGregorianTime(part1, part2 float64) time.Time {
	part1I, part1F := math.Modf(part1)
	part2I, part2F := math.Modf(part2)
	julianDays := part1I + part2I
	julianFraction := part1F + part2F
	julianDays, julianFraction = shiftJulianToNoon(julianDays, julianFraction)
	day, month, year := doTheFliegelAndVanFlandernAlgorithm(int(julianDays))
	hours, minutes, seconds, nanoseconds := fractionOfADay(julianFraction)
	return time.Date(year, time.Month(month), day, hours, minutes, seconds, nanoseconds, time.UTC)
}

// shiftJulianToNoon provides a function to process julian date to noon.
func shiftJulianToNoon(julianDays, julianFraction float64) (float64, float64) {
	switch {
	case -0.5 < julianFraction && julianFraction < 0.5:
		julianFraction += 0.5
	case julianFraction >= 0.5:
		julianDays++
		julianFraction -= 0.5
	case julianFraction <= -0.5:
		julianDays--
		julianFraction += 1.5
	}
	return julianDays, julianFraction
}
func doTheFliegelAndVanFlandernAlgorithm(jd int) (day, month, year int) {
	l := jd + 68569
	n := (4 * l) / 146097
	l = l - (146097*n+3)/4
	i := (4000 * (l + 1)) / 1461001
	l = l - (1461*i)/4 + 31
	j := (80 * l) / 2447
	d := l - (2447*j)/80
	l = j / 11
	m := j + 2 - (12 * l)
	y := 100*(n-49) + i + l
	return d, m, y
}

// fractionOfADay provides a function to return the integer values for hour,
// minutes, seconds and nanoseconds that comprised a given fraction of a day.
// values would round to 1 us.
func fractionOfADay(fraction float64) (hours, minutes, seconds, nanoseconds int) {

	const (
		c1us  = 1e3
		c1s   = 1e9
		c1day = 24 * 60 * 60 * c1s
	)

	frac := int64(c1day*fraction + c1us/2)
	nanoseconds = int((frac%c1s)/c1us) * c1us
	frac /= c1s
	seconds = int(frac % 60)
	frac /= 60
	minutes = int(frac % 60)
	hours = int(frac / 60)
	return
}

//Excel float value to time
func GetExcelTime(excelTime float64, date1904 bool) time.Time {
	const MDD int64 = 106750 // Max time.Duration Days, aprox. 290 years
	var date time.Time
	var intPart = int64(excelTime)
	// Excel uses Julian dates prior to March 1st 1900, and Gregorian
	// thereafter.
	if intPart <= 61 {
		const OFFSET1900 = 15018.0
		const OFFSET1904 = 16480.0
		const MJD0 float64 = 2400000.5
		var date time.Time
		if date1904 {
			date = julianDateToGregorianTime(MJD0, excelTime+OFFSET1904)
		} else {
			date = julianDateToGregorianTime(MJD0, excelTime+OFFSET1900)
		}
		return date
	}
	var floatPart = excelTime - float64(intPart)
	var dayNanoSeconds float64 = 24 * 60 * 60 * 1000 * 1000 * 1000
	if date1904 {
		date = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	} else {
		date = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	}

	// Duration is limited to aprox. 290 years
	for intPart > MDD {
		durationDays := time.Duration(MDD) * time.Hour * 24
		date = date.Add(durationDays)
		intPart = intPart - MDD
	}
	durationDays := time.Duration(intPart) * time.Hour * 24
	durationPart := time.Duration(dayNanoSeconds * floatPart)
	return date.Add(durationDays).Add(durationPart)
}
