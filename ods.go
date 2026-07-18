// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

// Package ods builds OpenDocument spreadsheet documents, both as zipped
// packages (.ods) and as flat XML documents (.fods).
//
// Cells are created with [MakeCell] or [MakeRangeCell], arranged in rows,
// and combined into a [Spreadsheet] with [MakeSpreadsheet]. The spreadsheet
// is then serialized with [MakeOds], [WriteOds], or [MakeFlatOds].
package ods

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"regexp"
	"strconv"
	"strings"
)

// generator identifies this library in the document metadata (meta:generator).
const generator = "github.com/fwilhe2/rechenbrett"

// odfVersion is the OpenDocument format version of the produced documents.
const odfVersion = "1.4"

// defaultTableName is the name of the single sheet created by MakeSpreadsheet.
const defaultTableName = "Sheet1"

var (
	germanDateFormat = regexp.MustCompile(`^(\d{1,2})\.(\d{1,2})\.(\d{4})$`)
	usDateFormat     = regexp.MustCompile(`^(\d{1,2})/(\d{1,2})/(\d{4})$`)
	isoDateFormat    = regexp.MustCompile(`^(\d{4})-(\d{1,2})-(\d{1,2})$`)
	timeFormat       = regexp.MustCompile(`^(\d{1,2}):(\d{2})(?::(\d{2}))?$`)
)

// MakeCell creates a cell holding value interpreted as valueType.
//
// Supported value types are "string", "float", "date" (ISO, German, or US
// format), "time" (HH:MM or HH:MM:SS), "percentage" (fraction, e.g. "0.42"
// for 42 %), "formula", and "currency" with the variants "currency-eur",
// "currency-usd", and "currency-gbp" (bare "currency" means EUR).
//
// Invalid values or value types are reported by [MakeSpreadsheet].
func MakeCell(value, valueType string) Cell {
	return createCell(cellData{
		Value:     value,
		ValueType: valueType,
	})
}

// MakeRangeCell creates a cell like [MakeCell] and additionally names its
// position as a range, so formulas in other cells can refer to it by
// rangeName. Each range name may be used for only one cell.
func MakeRangeCell(value, valueType, rangeName string) Cell {
	return createCell(cellData{
		Value:     value,
		ValueType: valueType,
		Range:     rangeName,
	})
}

// MakeSpreadsheet arranges the given rows of cells into a spreadsheet with a
// single sheet named "Sheet1".
//
// It reports all invalid cells (bad value types, unparseable dates, times, or
// numbers) and duplicate range names as a single joined error.
func MakeSpreadsheet(cells [][]Cell) (Spreadsheet, error) {
	return MakeSpreadsheetWithName(defaultTableName, cells)
}

// MakeSpreadsheetWithName is like [MakeSpreadsheet] with a custom sheet name.
func MakeSpreadsheetWithName(name string, cells [][]Cell) (Spreadsheet, error) {
	var rows []row
	var errs []error

	rangeAddresses := map[string]string{}
	rangeNames := []string{}

	maxCols := 1
	for rowIdx, c := range cells {
		rows = append(rows, row{Cells: c})
		maxCols = max(maxCols, len(c))
		for colIdx, cc := range c {
			if cc.err != nil {
				errs = append(errs, fmt.Errorf("row %d, column %d: %w", rowIdx+1, colIdx+1, cc.err))
			}
			if cc.rangeName == "" {
				continue
			}
			if _, exists := rangeAddresses[cc.rangeName]; exists {
				errs = append(errs, fmt.Errorf("row %d, column %d: duplicate range name %q", rowIdx+1, colIdx+1, cc.rangeName))
				continue
			}
			rangeAddresses[cc.rangeName] = fmt.Sprintf("$%s.%s", name, toA1(rowIdx+1, colIdx+1))
			rangeNames = append(rangeNames, cc.rangeName)
		}
	}
	if len(errs) > 0 {
		return Spreadsheet{}, errors.Join(errs...)
	}

	namedRanges := []namedRange{}
	for _, rangeName := range rangeNames {
		namedRanges = append(namedRanges, namedRange{
			Name:             rangeName,
			BaseCellAddress:  rangeAddresses[rangeName],
			CellRangeAddress: rangeAddresses[rangeName],
		})
	}

	return Spreadsheet{
		Tables: []table{
			{
				Name: name,
				// The ODF schema requires at least one table:table-column
				// before the table rows.
				Columns: []tableColumn{{NumberColumnsRepeated: strconv.Itoa(maxCols)}},
				Rows:    rows,
			},
		},
		NamedExpressions: namedExpressions{NamedRanges: namedRanges},
	}, nil
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

// MakeFlatOds serializes the spreadsheet as a flat OpenDocument XML document
// (.fods).
func MakeFlatOds(spreadsheet Spreadsheet) (string, error) {
	fods := flatOds{
		XMLNSOffice:    "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:     "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:      "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:     "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:    "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSMeta:      "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		OfficeVersion:  odfVersion,
		OfficeMimetype: "application/vnd.oasis.opendocument.spreadsheet",
		Meta:           officeMeta{Generator: generator},
		AutomaticStyles: automaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
		Body: documentBody{
			Spreadsheet: spreadsheet,
		},
	}

	out, err := xml.MarshalIndent(fods, " ", "  ")
	if err != nil {
		return "", fmt.Errorf("marshaling flat ods document: %w", err)
	}
	return xmlWithHeader(out), nil
}

// MakeOds serializes the spreadsheet as a zipped OpenDocument package (.ods).
func MakeOds(spreadsheet Spreadsheet) (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	if err := WriteOds(buf, spreadsheet); err != nil {
		return nil, err
	}
	return buf, nil
}

// WriteOds writes the spreadsheet as a zipped OpenDocument package (.ods)
// to w.
func WriteOds(w io.Writer, spreadsheet Spreadsheet) error {
	manifestXml := manifest{
		Version: odfVersion,
		XMLNS:   "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
		Entries: []fileEntry{
			{
				FullPath:  "/",
				Version:   odfVersion,
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
			{
				FullPath:  "meta.xml",
				MediaType: "text/xml",
			},
		},
	}

	contentXml := documentContent{
		XMLNSOffice:   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		OfficeVersion: odfVersion,
		AutomaticStyles: automaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
		Body: documentBody{
			Spreadsheet: spreadsheet,
		},
	}

	stylesXml := documentStyles{
		XMLNSOffice:   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSSvg:      "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		OfficeVersion: odfVersion,
		AutomaticStyles: automaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createStyles(),
		},
	}

	metaXml := documentMeta{
		XMLNSOffice:   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSMeta:     "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		OfficeVersion: odfVersion,
		Meta:          officeMeta{Generator: generator},
	}

	zipWriter := zip.NewWriter(w)

	// The mimetype entry must be stored uncompressed with its size and CRC in
	// the local file header (no data descriptor), or LibreOffice >= 26.2
	// refuses to load the file. CreateRaw writes the header as given, unlike
	// CreateHeader, which always streams with a data descriptor.
	mimetype := []byte("application/vnd.oasis.opendocument.spreadsheet")
	mimetypeHeader := &zip.FileHeader{
		Name:               "mimetype",
		Method:             zip.Store, // IMPORTANT: Use zip.Store for no compression.
		CRC32:              crc32.ChecksumIEEE(mimetype),
		CompressedSize64:   uint64(len(mimetype)),
		UncompressedSize64: uint64(len(mimetype)),
		// CreateRaw does not derive the MS-DOS timestamp from Modified, and a
		// zero timestamp (before 1980) makes zip readers like Apache Commons
		// Compress synthesize a timestamp extra field, which the ODF spec
		// forbids for the mimetype entry.
		ModifiedDate: (2026-1980)<<9 | 1<<5 | 1, // 2026-01-01
	}
	writer, err := zipWriter.CreateRaw(mimetypeHeader)
	if err != nil {
		return fmt.Errorf("creating mimetype zip entry: %w", err)
	}
	if _, err := writer.Write(mimetype); err != nil {
		return fmt.Errorf("writing mimetype zip entry: %w", err)
	}

	parts := []struct {
		name    string
		content any
	}{
		{"META-INF/manifest.xml", manifestXml},
		{"content.xml", contentXml},
		{"styles.xml", stylesXml},
		{"meta.xml", metaXml},
	}
	for _, part := range parts {
		marshaled, err := xml.MarshalIndent(part.content, "", "  ")
		if err != nil {
			return fmt.Errorf("marshaling %s: %w", part.name, err)
		}
		writer, err := zipWriter.Create(part.name)
		if err != nil {
			return fmt.Errorf("creating zip entry %s: %w", part.name, err)
		}
		if _, err := io.WriteString(writer, xmlWithHeader(marshaled)); err != nil {
			return fmt.Errorf("writing zip entry %s: %w", part.name, err)
		}
	}

	if err := zipWriter.Close(); err != nil {
		return fmt.Errorf("closing zip archive: %w", err)
	}

	return nil
}

func xmlWithHeader(input []byte) string {
	return xml.Header + string(input)
}

// timeString converts "HH:MM" or "HH:MM:SS" to an ISO 8601 duration as used
// by office:time-value.
func timeString(input string) (string, error) {
	matches := timeFormat.FindStringSubmatch(input)
	if matches == nil {
		return "", fmt.Errorf("invalid time %q, expected HH:MM or HH:MM:SS", input)
	}
	seconds := matches[3]
	if seconds == "" {
		seconds = "00"
	}
	return fmt.Sprintf("PT%sH%sM%sS", matches[1], matches[2], seconds), nil
}

// dateString converts German (DD.MM.YYYY), US (MM/DD/YYYY), or ISO
// (YYYY-MM-DD) dates to the ISO format used by office:date-value.
func dateString(date string) (string, error) {
	switch {
	case germanDateFormat.MatchString(date):
		matches := germanDateFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[3], matches[2], matches[1]), nil
	case usDateFormat.MatchString(date):
		matches := usDateFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[3], matches[1], matches[2]), nil
	case isoDateFormat.MatchString(date):
		matches := isoDateFormat.FindStringSubmatch(date)
		return fmt.Sprintf("%s-%02s-%02s", matches[1], matches[2], matches[3]), nil
	default:
		return "", fmt.Errorf("invalid date %q, expected DD.MM.YYYY, MM/DD/YYYY, or YYYY-MM-DD", date)
	}
}

func parseNumber(value, valueType string) error {
	if _, err := strconv.ParseFloat(value, 64); err != nil {
		return fmt.Errorf("invalid %s value %q, expected a number", valueType, value)
	}
	return nil
}

func createCell(data cellData) Cell {
	cell := Cell{
		ValueType: data.ValueType,
		rangeName: data.Range,
	}

	switch data.ValueType {
	case "string":
		cell.Text = data.Value
	case "float":
		cell.err = parseNumber(data.Value, data.ValueType)
		cell.StyleName = "FLOAT_STYLE"
		cell.Value = data.Value
	case "date":
		cell.StyleName = "DATE_STYLE"
		cell.DateValue, cell.err = dateString(data.Value)
	case "time":
		// No data style: LibreOffice formats time values with the locale's
		// default time format.
		cell.TimeValue, cell.err = timeString(data.Value)
	case "percentage":
		// No data style: LibreOffice formats percentage values with the
		// locale's default percentage format.
		cell.err = parseNumber(data.Value, data.ValueType)
		cell.Value = data.Value
	case "formula":
		cell.Formula = data.Value
		cell.ValueType = ""
	case "currency", "currency-eur", "currency-usd", "currency-gbp":
		// office:value-type only allows "currency"; the concrete currency is
		// given by office:currency. Bare "currency" defaults to EUR.
		cell.ValueType = "currency"
		cell.err = parseNumber(data.Value, data.ValueType)
		cell.Value = data.Value
		switch data.ValueType {
		case "currency-usd":
			cell.StyleName = "USD_STYLE"
			cell.Currency = "USD"
		case "currency-gbp":
			cell.StyleName = "GBP_STYLE"
			cell.Currency = "GBP"
		default:
			cell.StyleName = "EUR_STYLE"
			cell.Currency = "EUR"
		}
	default:
		cell.err = fmt.Errorf("unknown value type %q", data.ValueType)
	}
	return cell
}

// Data style names: the plain name renders negative values (red, with a
// minus sign) and maps to the _POSITIVE variant for values >= 0.
func createNumberStyles() []any {
	return []any{
		numberStyle{
			Name:     "FLOAT_DATA_STYLE_POSITIVE",
			Volatile: "true",
			NumberElements: []numberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
		},
		numberStyle{
			Name:           "FLOAT_DATA_STYLE",
			TextProperties: &textProperties{Color: "#ff0000"},
			Text:           "−",
			NumberElements: []numberElement{
				{
					DecimalPlaces:    "2",
					MinDecimalPlaces: "2",
					MinIntegerDigits: "1",
					Grouping:         "true",
				},
			},
			Map: &styleMap{Condition: "value()>=0", ApplyStyleName: "FLOAT_DATA_STYLE_POSITIVE"},
		},
		dateStyle{
			Name: "DATE_DATA_STYLE",
			Parts: []any{
				dateYear{Style: "long"},
				textElement{Content: "-"},
				dateMonth{Style: "long"},
				textElement{Content: "-"},
				dateDay{Style: "long"},
			},
		},
		makeCurrencyStyle("EUR", "de", "DE", "€"),
		makeNegativeCurrencyStyle("EUR", "de", "DE", "€"),
		makeCurrencyStyle("USD", "en", "US", "$"),
		makeNegativeCurrencyStyle("USD", "en", "US", "$"),
		makeCurrencyStyle("GBP", "en", "GB", "£"),
		makeNegativeCurrencyStyle("GBP", "en", "GB", "£"),
	}
}

func makeCurrencyStyle(code, language, country, symbol string) currencyStyle {
	return currencyStyle{
		Name:     code + "_DATA_STYLE_POSITIVE",
		Volatile: "true",
		Language: "en",
		Country:  country,
		Number: numberFormat{
			DecimalPlaces:    2,
			MinDecimalPlaces: 2,
			MinIntegerDigits: 1,
			Grouping:         true,
		},
		CurrencySymbol: currencySymbol{
			Language: language,
			Country:  country,
			Symbol:   symbol,
		},
	}
}

func makeNegativeCurrencyStyle(code, language, country, symbol string) currencyStyle {
	style := makeCurrencyStyle(code, language, country, symbol)
	style.Name = code + "_DATA_STYLE"
	style.TextProperties = &textProperties{Color: "#ff0000"}
	style.Texts = []textElement{{"−"}}
	style.StyleMap = &styleMap{Condition: "value()>=0", ApplyStyleName: code + "_DATA_STYLE_POSITIVE"}
	return style
}

func createStyles() []cellStyle {
	return []cellStyle{
		{Name: "FLOAT_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "FLOAT_DATA_STYLE"},
		{Name: "DATE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "DATE_DATA_STYLE"},
		{Name: "EUR_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "EUR_DATA_STYLE"},
		{Name: "USD_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "USD_DATA_STYLE"},
		{Name: "GBP_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "GBP_DATA_STYLE"},
	}
}

// Cell is one spreadsheet cell. Create cells with [MakeCell] or
// [MakeRangeCell].
type Cell struct {
	XMLName   xml.Name `xml:"table:table-cell"`
	Text      string   `xml:"text:p,omitempty"`
	ValueType string   `xml:"office:value-type,attr,omitempty"`
	Value     string   `xml:"office:value,attr,omitempty"`
	DateValue string   `xml:"office:date-value,attr,omitempty"`
	TimeValue string   `xml:"office:time-value,attr,omitempty"`
	Currency  string   `xml:"office:currency,attr,omitempty"`
	StyleName string   `xml:"table:style-name,attr,omitempty"`
	Formula   string   `xml:"table:formula,attr,omitempty"`

	rangeName string
	err       error
}

// Spreadsheet is a collection of tables ready for serialization. Create
// spreadsheets with [MakeSpreadsheet]; the fields are exported only for XML
// marshaling.
type Spreadsheet struct {
	XMLName          xml.Name         `xml:"office:spreadsheet"`
	Tables           []table          `xml:"table:table"`
	NamedExpressions namedExpressions `xml:"table:named-expressions"`
}

// cellData is the raw input for a cell before validation.
type cellData struct {
	Value     string
	ValueType string
	Range     string
}

type row struct {
	XMLName xml.Name `xml:"table:table-row"`
	Cells   []Cell   `xml:"table:table-cell"`
}

type tableColumn struct {
	XMLName               xml.Name `xml:"table:table-column"`
	NumberColumnsRepeated string   `xml:"table:number-columns-repeated,attr,omitempty"`
}

type table struct {
	XMLName xml.Name      `xml:"table:table"`
	Name    string        `xml:"table:name,attr"`
	Columns []tableColumn `xml:"table:table-column"`
	Rows    []row         `xml:"table:table-row"`
}

type flatOds struct {
	XMLName         xml.Name        `xml:"office:document"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSMeta       string          `xml:"xmlns:meta,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	OfficeMimetype  string          `xml:"office:mimetype,attr"`
	Meta            officeMeta      `xml:"office:meta"`
	AutomaticStyles automaticStyles `xml:"office:automatic-styles"`
	Body            documentBody    `xml:"office:body"`
}

type documentContent struct {
	XMLName         xml.Name        `xml:"office:document-content"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	AutomaticStyles automaticStyles `xml:"office:automatic-styles"`
	Body            documentBody    `xml:"office:body"`
}

type documentMeta struct {
	XMLName       xml.Name   `xml:"office:document-meta"`
	XMLNSOffice   string     `xml:"xmlns:office,attr"`
	XMLNSMeta     string     `xml:"xmlns:meta,attr"`
	OfficeVersion string     `xml:"office:version,attr"`
	Meta          officeMeta `xml:"office:meta"`
}

type officeMeta struct {
	XMLName   xml.Name `xml:"office:meta"`
	Generator string   `xml:"meta:generator"`
}

type documentStyles struct {
	XMLName         xml.Name        `xml:"office:document-styles"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSSvg        string          `xml:"xmlns:svg,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	AutomaticStyles automaticStyles `xml:"office:automatic-styles"`
}

type automaticStyles struct {
	XMLName      xml.Name    `xml:"office:automatic-styles"`
	NumberStyles []any       `xml:"number:number-style"`
	Styles       []cellStyle `xml:"style:style"`
}

// Field order matters: the ODF schema requires style:text-properties to be
// the first child of number:number-style, before any number:text.
type numberStyle struct {
	XMLName        xml.Name        `xml:"number:number-style"`
	Name           string          `xml:"style:name,attr"`
	Volatile       string          `xml:"style:volatile,attr,omitempty"`
	Language       string          `xml:"number:language,attr,omitempty"`
	Country        string          `xml:"number:country,attr,omitempty"`
	TextProperties *textProperties `xml:"style:text-properties,omitempty"`
	Text           string          `xml:"number:text,omitempty"`
	NumberElements []numberElement `xml:",any"`
	Map            *styleMap       `xml:"style:map,omitempty"`
}

type textProperties struct {
	Color string `xml:"fo:color,attr,omitempty"`
}

type numberElement struct {
	XMLName          xml.Name `xml:"number:number"`
	DecimalPlaces    string   `xml:"number:decimal-places,attr,omitempty"`
	MinDecimalPlaces string   `xml:"number:min-decimal-places,attr,omitempty"`
	MinIntegerDigits string   `xml:"number:min-integer-digits,attr,omitempty"`
	Grouping         string   `xml:"number:grouping,attr,omitempty"`
}

type styleMap struct {
	XMLName        xml.Name `xml:"style:map"`
	Condition      string   `xml:"style:condition,attr"`
	ApplyStyleName string   `xml:"style:apply-style-name,attr"`
}

type cellStyle struct {
	XMLName         xml.Name `xml:"style:style"`
	Name            string   `xml:"style:name,attr"`
	Family          string   `xml:"style:family,attr"`
	ParentStyleName string   `xml:"style:parent-style-name,attr"`
	DataStyleName   string   `xml:"style:data-style-name,attr"`
}

type documentBody struct {
	XMLName     xml.Name    `xml:"office:body"`
	Spreadsheet Spreadsheet `xml:"office:spreadsheet"`
}

type manifest struct {
	XMLName xml.Name    `xml:"manifest:manifest"`
	Version string      `xml:"manifest:version,attr"`
	XMLNS   string      `xml:"xmlns:manifest,attr"`
	Entries []fileEntry `xml:"manifest:file-entry"`
}

type fileEntry struct {
	FullPath  string `xml:"manifest:full-path,attr"`
	Version   string `xml:"manifest:version,attr,omitempty"`
	MediaType string `xml:"manifest:media-type,attr"`
}

// Field order matters: the ODF schema requires style:text-properties to be
// the first child of number:currency-style.
type currencyStyle struct {
	XMLName        xml.Name        `xml:"number:currency-style"`
	Name           string          `xml:"style:name,attr"`
	Volatile       string          `xml:"style:volatile,attr,omitempty"`
	Language       string          `xml:"number:language,attr"`
	Country        string          `xml:"number:country,attr"`
	TextProperties *textProperties `xml:"style:text-properties,omitempty"`
	Texts          []textElement   `xml:"number:text"`
	Number         numberFormat    `xml:"number:number"`
	CurrencySymbol currencySymbol  `xml:"number:currency-symbol"`
	StyleMap       *styleMap       `xml:"style:map,omitempty"`
}

type numberFormat struct {
	DecimalPlaces    int  `xml:"number:decimal-places,attr"`
	MinDecimalPlaces int  `xml:"number:min-decimal-places,attr"`
	MinIntegerDigits int  `xml:"number:min-integer-digits,attr"`
	Grouping         bool `xml:"number:grouping,attr"`
}

type textElement struct {
	Content string `xml:",chardata"`
}

type currencySymbol struct {
	Language string `xml:"number:language,attr"`
	Country  string `xml:"number:country,attr"`
	Symbol   string `xml:",chardata"`
}

type dateStyle struct {
	XMLName xml.Name `xml:"number:date-style"`
	Name    string   `xml:"style:name,attr"`
	Parts   []any    `xml:"number:text"`
}

type dateYear struct {
	XMLName xml.Name `xml:"number:year"`
	Style   string   `xml:"number:style,attr"`
}

type dateMonth struct {
	XMLName xml.Name `xml:"number:month"`
	Style   string   `xml:"number:style,attr"`
}

type dateDay struct {
	XMLName xml.Name `xml:"number:day"`
	Style   string   `xml:"number:style,attr"`
}

type namedExpressions struct {
	NamedRanges []namedRange `xml:"table:named-range"`
}

type namedRange struct {
	Name             string `xml:"table:name,attr"`
	BaseCellAddress  string `xml:"table:base-cell-address,attr"`
	CellRangeAddress string `xml:"table:cell-range-address,attr"`
}
