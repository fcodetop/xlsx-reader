package xlsx_reader

import (
	"encoding/csv"
	"errors"
	"github.com/axgle/mahonia"
	"golang.org/x/net/html/charset"
	"os"
	"strings"
)

func ReadCsv(filePath string, encodeing string, cols *[]string, totalCount *int, rowAction func(row []string)) error {
	f, err := os.Open(filePath)
	if err != nil {
		return err
	}
	defer f.Close()
	var reader *csv.Reader
	if encodeing == "" {
		//auto detect csv file charset
		content := make([]byte, 128, 128)
		_, err = f.Read(content)
		if err != nil {
			return err
		}
		_, encodeing, _ = charset.DetermineEncoding(content, "text/csv")
		f.Seek(0, 0)

	}

	decoder := mahonia.NewDecoder(encodeing)
	reader = csv.NewReader(decoder.NewReader(f))

	reader.Comment = '#'

	data, err := reader.ReadAll()
	if err != nil {
		return err
	}
	r := len(data) - 1
	*totalCount = r
	if r > 0 {
		data[0][0] = strings.ReplaceAll(data[0][0], "\ufeff", "") //utf8-bom
		m, err := checkCol1(cols, &data[0])
		if err != nil {
			return err
		}
		c := len(*cols)
		row := make([]string, c)
		for i := 1; i <= r; i++ {
			for j := 0; j < c; j++ {
				row[j] = data[i][m[j]]
			}
			rowAction(row)
		}
	}
	return nil
}
func checkCol1(cols *[]string, header *[]string) ([]int, error) {
	l := len(*cols)
	if l > len(*header) {
		return nil, errors.New(matchColsErr)
	}
	m := make([]int, l)
	f := -1
	for i, v := range *cols {
		for j, c := range *header {
			if v == c {
				m[i] = j
				if j == 0 {
					f = i
				}
				break
			}
		}
		if m[i] == 0 && f != i {
			return nil, errors.New(matchColsErr)
		}
	}
	return m, nil
}
