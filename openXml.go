package xlsx_reader

import "encoding/xml"

type xlsxWorkbook struct {
	XMLName xml.Name   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	Sheets  xlsxSheets `xml:"sheets"`
}

// xlsxSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main.
type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

// xlsxSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxSheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetID string `xml:"sheetId,attr,omitempty"`
	ID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	//State   string `xml:"state,attr,omitempty"`
}

// xmlxWorkbookRels contains xmlxWorkbookRelations which maps sheet id and sheet XML.
type xlsxWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorkbookRelation `xml:"Relationship"`
}

// xmlxWorkbookRelation maps sheet id and xl/worksheets/_rels/sheet%d.xml.rels
type xlsxWorkbookRelation struct {
	ID     string `xml:"Id,attr"`
	Target string `xml:",attr"`
	//Type       string `xml:",attr"`
	//TargetMode string `xml:",attr,omitempty"`
}

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxWorksheet struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	//SheetPr               *xlsxSheetPr                 `xml:"sheetPr"`
	//Dimension             xlsxDimension                `xml:"dimension"`
	//SheetViews            xlsxSheetViews               `xml:"sheetViews,omitempty"`
	//SheetFormatPr         *xlsxSheetFormatPr           `xml:"sheetFormatPr"`
	//Cols                  *xlsxCols                    `xml:"cols,omitempty"`
	SheetData xlsxSheetData `xml:"sheetData"`
	//SheetProtection       *xlsxSheetProtection         `xml:"sheetProtection"`
	//AutoFilter            *xlsxAutoFilter              `xml:"autoFilter"`
	//MergeCells            *xlsxMergeCells              `xml:"mergeCells"`
	//PhoneticPr            *xlsxPhoneticPr              `xml:"phoneticPr"`
	//ConditionalFormatting []*xlsxConditionalFormatting `xml:"conditionalFormatting"`
	//DataValidations       *xlsxDataValidations         `xml:"dataValidations,omitempty"`
	//Hyperlinks            *xlsxHyperlinks              `xml:"hyperlinks"`
	//PrintOptions          *xlsxPrintOptions            `xml:"printOptions"`
	//PageMargins           *xlsxPageMargins             `xml:"pageMargins"`
	//PageSetUp             *xlsxPageSetUp               `xml:"pageSetup"`
	//HeaderFooter          *xlsxHeaderFooter            `xml:"headerFooter"`
	//Drawing               *xlsxDrawing                 `xml:"drawing"`
	//LegacyDrawing         *xlsxLegacyDrawing           `xml:"legacyDrawing"`
	//Picture               *xlsxPicture                 `xml:"picture"`
	//TableParts            *xlsxTableParts              `xml:"tableParts"`
	//ExtLst                *xlsxExtLst                  `xml:"extLst"`
}

// xlsxSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxSheetData struct {
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

// xlsxRow directly maps the row element. The element expresses information
// about an entire row of a worksheet, and contains all cell definitions for a
// particular row in the worksheet.
type xlsxRow struct {
	//Collapsed    bool    `xml:"collapsed,attr,omitempty"`
	//CustomFormat bool    `xml:"customFormat,attr,omitempty"`
	//CustomHeight bool    `xml:"customHeight,attr,omitempty"`
	//Hidden       bool    `xml:"hidden,attr,omitempty"`
	//Ht           float64 `xml:"ht,attr,omitempty"`
	//OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
	//Ph           bool    `xml:"ph,attr,omitempty"`
	//R            int     `xml:"r,attr,omitempty"`
	//S            int     `xml:"s,attr,omitempty"`
	//Spans        string  `xml:"spans,attr,omitempty"`
	//ThickBot     bool    `xml:"thickBot,attr,omitempty"`
	//ThickTop     bool    `xml:"thickTop,attr,omitempty"`
	C []xlsxC `xml:"c"`
}

//xlsxC directly maps the cell element.
type xlsxC struct {
	R string `xml:"r,attr"` // Cell ID, e.g. A1
	//S int    `xml:"s,attr,omitempty"` // Style reference.
	T string `xml:"t,attr,omitempty"` // Type.
	//F        *xlsxF   `xml:"f,omitempty"`      // Formula
	V string `xml:"v,omitempty"` // Value
	//IS       *xlsxIS  `xml:"is"`
	//XMLSpace xml.Attr `xml:"space,attr,omitempty"`
}

// xlsxSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main. String values may
// be stored directly inside spreadsheet cell elements; however, storing the
// same value inside multiple cell elements can result in very large worksheet
// Parts, possibly resulting in kpis degradation. The Shared String Table
// is an indexed list of string values, shared across the workbook, which allows
// implementations to store values only once.
type xlsxSST struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	//Count       int      `xml:"count,attr"`
	//UniqueCount int      `xml:"uniqueCount,attr"`
	SI []xlsxSI `xml:"si"`
}

// xlsxSI directly maps the si element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxSI struct {
	T string `xml:"t"`
}
