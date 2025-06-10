// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"log"
	"regexp"
	"strings"
)

func MakeCell(value, valueType string) Cell {
	return createCell(CellData{
		Value:     value,
		ValueType: valueType,
	})
}

func MakeRangeCell(value, valueType, rangeName string) Cell {
	return createCell(CellData{
		Value:     value,
		ValueType: valueType,
		Range:     rangeName,
	})
}

func MakeSpreadsheet(cells [][]Cell) Spreadsheet {
	var rows []Row

	rangesData := map[string][][2]int{}
	ranges := []string{}

	r1 := 1
	c1 := 1
	for _, c := range cells {
		rows = append(rows, Row{Cells: c})
		for _, cc := range c {
			if len(cc.Range) > 0 {
				rangesData[cc.Range] = append(rangesData[cc.Range], [2]int{r1, c1})
				ranges = append(ranges, cc.Range)
			}
			c1++
		}
		r1++
		c1 = 1
	}

	tables := []Table{
		{
			Name: "Sheet1",
			Rows: rows,
		},
	}

	namedRanges := []NamedRange{}
	for rangeIdx := range ranges {
		namedRanges = append(namedRanges, NamedRange{
			Name:             ranges[rangeIdx],
			BaseCellAddress:  fmt.Sprintf("$Sheet1.%s", toA1(rangesData[ranges[rangeIdx]][0][0], rangesData[ranges[rangeIdx]][0][1])),
			CellRangeAddress: fmt.Sprintf("$Sheet1.%s", toA1(rangesData[ranges[rangeIdx]][0][0], rangesData[ranges[rangeIdx]][0][1])),
		})
	}

	return Spreadsheet{
		Tables:           tables,
		NamedExpressions: NamedExpressions{NamedRanges: namedRanges},
	}
}

// Converts a column number to its Excel-style letter representation
func columnToLetters(col int) string {
	letters := ""
	for col > 0 {
		col-- // adjust for 1-based indexing
		letter := 'A' + (col % 26)
		letters = string(rune(letter)) + letters
		col /= 26
	}
	return letters
}

// Converts row and column indices to Excel A1 format
func toA1(row, col int) string {
	if row < 1 || col < 1 {
		return ""
	}
	return fmt.Sprintf("$%s$%d", columnToLetters(col), row)
}

func MakeFlatOds(spreadsheet Spreadsheet) string {
	fods := FlatOds{
		XMLNSOffice:    "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:     "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:      "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:     "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:    "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSCalcext:   "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
		OfficeVersion:  "1.3",
		OfficeMimetype: "application/vnd.oasis.opendocument.spreadsheet",
		AutomaticStyles: AutomaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
		Body: Body{
			Spreadsheet: spreadsheet,
		},
	}

	out, _ := xml.MarshalIndent(fods, " ", "  ")
	return xmlByteArrayToStringWithHeader(out)
}

func MakeOds(spreadsheet Spreadsheet) *bytes.Buffer {
	manifest := Manifest{
		XMLNS: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
		Entries: []FileEntry{
			{
				FullPath:  "/",
				Version:   "1.3",
				MediaType: "application/vnd.oasis.opendocument.spreadsheet",
			},
			{
				FullPath:  "styles.xml",
				MediaType: "text/xml",
			},
			{
				FullPath:  "content.xml",
				MediaType: "text/xml",
			},
		},
	}

	contentXml := OfficeDocumentContent{
		XMLNSOffice:   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSCalcext:  "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
		OfficeVersion: "1.3",
		AutomaticStyles: AutomaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
		Body: Body{
			Spreadsheet: spreadsheet,
		},
	}

	stylesXml := OfficeDocumentStyles{
		XMLNSOffice:   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSSvg:      "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		XMLNSCalcext:  "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
		OfficeVersion: "1.3",
		AutomaticStyles: AutomaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
	}

	manifestStr, err := xml.MarshalIndent(manifest, "", "  ")
	if err != nil {
		log.Panic(err)
	}

	contentStr, err := xml.MarshalIndent(contentXml, "", "  ")
	if err != nil {
		log.Panic(err)
	}

	stylesStr, err := xml.MarshalIndent(stylesXml, "", "  ")
	if err != nil {
		log.Panic(err)
	}

	byesBuffer := new(bytes.Buffer)
	w := zip.NewWriter(byesBuffer)

	files := []struct {
		Name, Body string
	}{
		{"mimetype", "application/vnd.oasis.opendocument.spreadsheet"},
		{"META-INF/manifest.xml", xmlByteArrayToStringWithHeader(manifestStr)},
		{"content.xml", xmlByteArrayToStringWithHeader(contentStr)},
		{"styles.xml", xmlByteArrayToStringWithHeader(stylesStr)},
	}
	for _, file := range files {
		f, err := w.Create(file.Name)
		if err != nil {
			log.Fatal(err)
		}
		_, err = f.Write([]byte(file.Body))
		if err != nil {
			log.Fatal(err)
		}
	}

	// Make sure to check the error on Close.
	err = w.Close()
	if err != nil {
		log.Fatal(err)
	}

	return byesBuffer
}

func xmlByteArrayToStringWithHeader(input []byte) string {
	return xml.Header + string(input)
}

func timeString(input string) string {
	parts := strings.Split(input, ":")
	if len(parts) == 2 {
		return "PT" + parts[0] + "H" + parts[1] + "M00S"
	}
	if len(parts) == 3 {
		return "PT" + parts[0] + "H" + parts[1] + "M" + parts[2] + "S"
	}
	panic("Invalid time format " + input)
}

func dateString(date string) string {
	germanFormat := regexp.MustCompile(`^(\d{1,2})\.(\d{1,2})\.(\d{4})$`)
	usFormat := regexp.MustCompile(`^(\d{1,2})/(\d{1,2})/(\d{4})$`)
	isoFormat := regexp.MustCompile(`^(\d{4})-(\d{1,2})-(\d{1,2})$`)

	switch {
	case germanFormat.MatchString(date):
		matches := germanFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[3], matches[2], matches[1])
	case usFormat.MatchString(date):
		matches := usFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[3], matches[1], matches[2])
	case isoFormat.MatchString(date):
		matches := isoFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[1], matches[2], matches[3])
	default:
		panic("unknown date format")
	}
}

func createCell(cellData CellData) Cell {
	cell := Cell{
		ValueType: cellData.ValueType,
		Range:     cellData.Range,
	}
	switch cellData.ValueType {
	case "string":
		cell.Text = cellData.Value
		cell.CalcExtType = "string"
	case "float":
		cell.CalcExtType = "float"
		cell.StyleName = "FLOAT_STYLE"
		cell.Value = cellData.Value
	case "date":
		cell.CalcExtType = "date"
		cell.StyleName = "DATE_STYLE"
		cell.DateValue = dateString(cellData.Value)
	case "time":
		cell.CalcExtType = "time"
		cell.StyleName = "TIME_STYLE"
		cell.TimeValue = timeString(cellData.Value)
	case "currency":
		cell.CalcExtType = "currency"
		cell.StyleName = "EUR_STYLE"
		cell.Value = cellData.Value
		cell.Currency = "EUR"
	case "percentage":
		cell.CalcExtType = "percentage"
		cell.StyleName = "PERCENTAGE_STYLE"
		cell.Value = cellData.Value
	case "formula":
		cell.Formula = cellData.Value
		cell.ValueType = ""
	}
	return cell
}

func createNumberStyles() []interface{} {
	return []interface{}{
		NumberStyle{
			Name:     "___FLOAT_STYLE",
			Volatile: "true",
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
		},
		NumberStyle{
			Name:           "__FLOAT_STYLE",
			TextProperties: &TextProperties{Color: "#ff0000"},
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
			Map: &Map{Condition: "value()>=0", ApplyStyleName: "___FLOAT_STYLE"},
		},
		DateStyle{
			Name: "__DATE_STYLE",
			Parts: []interface{}{
				NumberElementDateYear{XMLName: xml.Name{Local: "number:year"}, Style: "long"},
				TextElement{Content: "-"},
				NumberElementDateMonth{XMLName: xml.Name{Local: "number:month"}, Style: "long"},
				TextElement{Content: "-"},
				NumberElementDateDay{XMLName: xml.Name{Local: "number:day"}, Style: "long"},
			},
		},
		NumberStyle{
			Name: "__TIME_STYLE",
			NumberElements: []NumberElement{
				{XMLName: xml.Name{Local: "number:hours"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: ":"},
				{XMLName: xml.Name{Local: "number:minutes"}, DecimalPlaces: "long"},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: ":"},
				{XMLName: xml.Name{Local: "number:seconds"}, DecimalPlaces: "long"},
			},
		},
		CurrencyStyle{
			Name:     "__EUR_STYLE",
			Volatile: "true",
			Language: "en",
			Country:  "DE",
			Number: NumberFormat{
				DecimalPlaces:    2,
				MinIntegerDigits: 1,
				Grouping:         true,
			},
			Texts: []TextElement{{}},
			CurrencySymbol: CurrencySymbol{
				Language: "de",
				Country:  "DE",
				Symbol:   "€",
			},
		},
		NumberStyle{
			Name: "__PERCENTAGE_STYLE",
			NumberElements: []NumberElement{
				{
					DecimalPlaces:    "2",
					MinIntegerDigits: "1",
				},
				{XMLName: xml.Name{Local: "number:text"}, DecimalPlaces: "%"},
			},
		},
	}
}

func createStyles() []Style {
	return []Style{
		{Name: "FLOAT_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__FLOAT_STYLE"},
		{Name: "DATE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__DATE_STYLE"},
		{Name: "TIME_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__TIME_STYLE"},
		{Name: "EUR_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__EUR_STYLE"},
		{Name: "PERCENTAGE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "__PERCENTAGE_STYLE"},
	}
}

type Cell struct {
	XMLName     xml.Name `xml:"table:table-cell"`
	Text        string   `xml:"text:p,omitempty"`
	ValueType   string   `xml:"office:value-type,attr,omitempty"`
	CalcExtType string   `xml:"calcext:value-type,attr,omitempty"`
	Value       string   `xml:"office:value,attr,omitempty"`
	DateValue   string   `xml:"office:date-value,attr,omitempty"`
	TimeValue   string   `xml:"office:time-value,attr,omitempty"`
	Currency    string   `xml:"office:currency,attr,omitempty"`
	StyleName   string   `xml:"table:style-name,attr,omitempty"`
	Formula     string   `xml:"table:formula,attr,omitempty"`
	Range       string   `xml:"-"`
}

type Row struct {
	XMLName xml.Name `xml:"table:table-row"`
	Cells   []Cell   `xml:"table:table-cell"`
}

type Table struct {
	XMLName xml.Name `xml:"table:table"`
	Name    string   `xml:"table:name,attr"`
	Rows    []Row    `xml:"table:table-row"`
}

type FlatOds struct {
	XMLName         xml.Name        `xml:"office:document"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSCalcext    string          `xml:"xmlns:calcext,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	OfficeMimetype  string          `xml:"office:mimetype,attr"`
	AutomaticStyles AutomaticStyles `xml:"office:automatic-styles"`
	Body            Body            `xml:"office:body"`
}

type OfficeDocumentContent struct {
	XMLName         xml.Name        `xml:"office:document-content"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSCalcext    string          `xml:"xmlns:calcext,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	AutomaticStyles AutomaticStyles `xml:"office:automatic-styles"`
	Body            Body            `xml:"office:body"`
}

type OfficeDocumentStyles struct {
	XMLName         xml.Name        `xml:"office:document-styles"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSSvg        string          `xml:"xmlns:svg,attr"`
	XMLNSCalcext    string          `xml:"xmlns:calcext,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	AutomaticStyles AutomaticStyles `xml:"office:automatic-styles"`
}

type Styles struct {
	XMLName xml.Name      `xml:"styles"`
	Items   []interface{} `xml:",any"`
}

type AutomaticStyles struct {
	XMLName      xml.Name      `xml:"office:automatic-styles"`
	NumberStyles []interface{} `xml:"number:number-style"`
	Styles       []Style       `xml:"style:style"`
}

type NumberStyle struct {
	XMLName        xml.Name        `xml:"number:number-style"`
	Name           string          `xml:"style:name,attr"`
	Volatile       string          `xml:"style:volatile,attr,omitempty"`
	Language       string          `xml:"number:language,attr,omitempty"`
	Country        string          `xml:"number:country,attr,omitempty"`
	TextProperties *TextProperties `xml:"style:text-properties,omitempty"`
	NumberElements []NumberElement `xml:",any"`
	Map            *Map            `xml:"style:map,omitempty"`
}

type TextProperties struct {
	XMLName xml.Name `xml:"style:text-properties"`
	Color   string   `xml:"fo:color,attr,omitempty"`
}

type NumberElement struct {
	XMLName          xml.Name `xml:"number:number"`
	DecimalPlaces    string   `xml:"number:decimal-places,attr"`
	MinDecimalPlaces string   `xml:"number:min-decimal-places,attr"`
	MinIntegerDigits string   `xml:"number:min-integer-digits,attr"`
	Grouping         string   `xml:"number:grouping,attr"`
	Language         string   `xml:"-"`
	Country          string   `xml:"-"`
}

type Map struct {
	XMLName        xml.Name `xml:"style:map"`
	Condition      string   `xml:"style:condition,attr"`
	ApplyStyleName string   `xml:"style:apply-style-name,attr"`
}

type Style struct {
	XMLName         xml.Name `xml:"style:style"`
	Name            string   `xml:"style:name,attr"`
	Family          string   `xml:"style:family,attr"`
	ParentStyleName string   `xml:"style:parent-style-name,attr"`
	DataStyleName   string   `xml:"style:data-style-name,attr"`
}

type Body struct {
	XMLName     xml.Name    `xml:"office:body"`
	Spreadsheet Spreadsheet `xml:"office:spreadsheet"`
}

type Spreadsheet struct {
	XMLName          xml.Name         `xml:"office:spreadsheet"`
	Tables           []Table          `xml:"table:table"`
	NamedExpressions NamedExpressions `xml:"table:named-expressions"`
}

type CellData struct {
	Value     string `json:"value"`
	ValueType string `json:"valueType"`
	Range     string `json:"range,omitempty"`
}

type Manifest struct {
	XMLName xml.Name    `xml:"manifest:manifest"`
	Version string      `xml:"manifest:version,attr"`
	XMLNS   string      `xml:"xmlns:manifest,attr"`
	Entries []FileEntry `xml:"manifest:file-entry"`
}

type FileEntry struct {
	FullPath  string `xml:"manifest:full-path,attr"`
	Version   string `xml:"manifest:version,attr,omitempty"`
	MediaType string `xml:"manifest:media-type,attr"`
}

type CurrencyStyle struct {
	XMLName        xml.Name        `xml:"number:currency-style"`
	Name           string          `xml:"style:name,attr"`
	Volatile       string          `xml:"style:volatile,attr,omitempty"`
	Language       string          `xml:"number:language,attr"`
	Country        string          `xml:"number:country,attr"`
	Number         NumberFormat    `xml:"number:number"`
	Texts          []TextElement   `xml:"number:text"`
	CurrencySymbol CurrencySymbol  `xml:"number:currency-symbol"`
	TextProperties *TextProperties `xml:"style:text-properties,omitempty"`
	StyleMap       *StyleMap       `xml:"style:map,omitempty"`
}

type NumberFormat struct {
	DecimalPlaces    int  `xml:"number:decimal-places,attr"`
	MinDecimalPlaces int  `xml:"number:min-decimal-places,attr"`
	MinIntegerDigits int  `xml:"number:min-integer-digits,attr"`
	Grouping         bool `xml:"number:grouping,attr"`
}

type TextElement struct {
	Content string `xml:",chardata"`
}

type CurrencySymbol struct {
	Language string `xml:"number:language,attr"`
	Country  string `xml:"number:country,attr"`
	Symbol   string `xml:",chardata"`
}

type StyleMap struct {
	Condition      string `xml:"style:condition,attr"`
	ApplyStyleName string `xml:"style:apply-style-name,attr"`
}

type DateStyle struct {
	XMLName xml.Name      `xml:"number:date-style"`
	Name    string        `xml:"style:name,attr"`
	Parts   []interface{} `xml:"number:text"`
}

type NumberElementDateYear struct {
	XMLName xml.Name `xml:"number:year"`
	Style   string   `xml:"number:style,attr"`
}

type NumberElementDateMonth struct {
	XMLName xml.Name `xml:"number:month"`
	Style   string   `xml:"number:style,attr"`
}

type NumberElementDateDay struct {
	XMLName xml.Name `xml:"number:day"`
	Style   string   `xml:"number:style,attr"`
}

type NamedExpressions struct {
	NamedRanges []NamedRange `xml:"table:named-range"`
}

type NamedRange struct {
	Name             string `xml:"table:name,attr"`
	BaseCellAddress  string `xml:"table:base-cell-address,attr"`
	CellRangeAddress string `xml:"table:cell-range-address,attr"`
}
