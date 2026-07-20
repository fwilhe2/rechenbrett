// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

// Package ods builds OpenDocument spreadsheet documents, both as zipped
// packages (.ods) and as flat XML documents (.fods).
//
// Cells are created with [MakeCell], [MakeRangeCell], or [MakeStyledCell],
// arranged in rows, and combined into a [Spreadsheet] with
// [MakeSpreadsheet]. The spreadsheet is then serialized with [MakeOds],
// [WriteOds], or [MakeFlatOds].
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
	"time"
)

// generator identifies this library in the document metadata (meta:generator).
const generator = "github.com/fwilhe2/rechenbrett"

// odfVersion is the OpenDocument format version of the produced documents.
const odfVersion = "1.4"

// defaultTableName is the name of the single sheet created by MakeSpreadsheet.
const defaultTableName = "Sheet1"

// zipEntryTime is the modification time stamped on every entry of a written
// package, so that the output is reproducible and the timestamp is
// representable in the MS-DOS date fields of a zip archive.
var zipEntryTime = time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC)

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

// CellStyle customizes the visual appearance of a cell. Colors are hex
// strings such as "#ff0000". Border, if set, is an ODF fo:border shorthand
// value (e.g. "0.5pt solid #000000") applied to all four sides of the cell.
// Zero-value fields are left unset.
type CellStyle struct {
	BackgroundColor string
	FontColor       string
	Bold            bool
	Italic          bool
	Border          string
}

// MakeStyledCell creates a cell like [MakeCell], additionally applying style
// to its appearance. Cells created with an identical style share a single
// generated style definition.
func MakeStyledCell(value, valueType string, style CellStyle) Cell {
	return createCell(cellData{
		Value:     value,
		ValueType: valueType,
		Style:     &style,
	})
}

// Color constants for use with [CellStyle.BackgroundColor] and
// [CellStyle.FontColor], taken from the palette at https://clrs.cc/.
const (
	ColorNavy    = "#001f3f"
	ColorBlue    = "#0074d9"
	ColorAqua    = "#7fdbff"
	ColorTeal    = "#39cccc"
	ColorPurple  = "#b10dc9"
	ColorFuchsia = "#f012be"
	ColorMaroon  = "#85144b"
	ColorRed     = "#ff4136"
	ColorOrange  = "#ff851b"
	ColorYellow  = "#ffdc00"
	ColorOlive   = "#3d9970"
	ColorGreen   = "#2ecc40"
	ColorLime    = "#01ff70"
	ColorBlack   = "#111111"
	ColorGray    = "#aaaaaa"
	ColorSilver  = "#dddddd"
	ColorWhite   = "#ffffff"
)

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

	customStyleNames := map[customStyleKey]string{}
	var customStyles []cellStyle

	maxCols := 1
	for rowIdx, c := range cells {
		rows = append(rows, row{Cells: c})
		maxCols = max(maxCols, len(c))
		for colIdx, cc := range c {
			if cc.err != nil {
				errs = append(errs, fmt.Errorf("row %d, column %d: %w", rowIdx+1, colIdx+1, cc.err))
			}
			if cc.rangeName != "" {
				if _, exists := rangeAddresses[cc.rangeName]; exists {
					errs = append(errs, fmt.Errorf("row %d, column %d: duplicate range name %q", rowIdx+1, colIdx+1, cc.rangeName))
				} else {
					rangeAddresses[cc.rangeName] = fmt.Sprintf("$%s.%s", name, toA1(rowIdx+1, colIdx+1))
					rangeNames = append(rangeNames, cc.rangeName)
				}
			}
			if cc.style != nil {
				key := customStyleKey{CellStyle: *cc.style, dataStyleName: dataStyleNameFor(cc.StyleName)}
				styleName, exists := customStyleNames[key]
				if !exists {
					styleName = fmt.Sprintf("CUSTOM_STYLE_%d", len(customStyleNames)+1)
					customStyleNames[key] = styleName
					customStyles = append(customStyles, buildCustomCellStyle(styleName, key.dataStyleName, *cc.style))
				}
				c[colIdx].StyleName = styleName
			}
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
				Name:      name,
				StyleName: tableStyleName,
				// The ODF schema requires at least one table:table-column
				// before the table rows.
				Columns: []tableColumn{{NumberColumnsRepeated: strconv.Itoa(maxCols)}},
				Rows:    rows,
			},
		},
		NamedExpressions: namedExpressions{NamedRanges: namedRanges},
		customStyles:     customStyles,
	}, nil
}

// EnableAutoFilter turns on AutoFilter dropdown buttons over the used cell
// range of every non-empty sheet in spreadsheet, so the generated document
// opens with filter dropdowns on each sheet's data. It emits a
// table:database-range with table:display-filter-buttons="true" but no saved
// filter conditions, leaving all rows visible.
//
// Calling it more than once replaces any previously enabled AutoFilter.
func EnableAutoFilter(spreadsheet Spreadsheet) Spreadsheet {
	var ranges []databaseRange
	for i, t := range spreadsheet.Tables {
		rowCount := len(t.Rows)
		if rowCount == 0 {
			continue
		}
		colCount := 1
		for _, r := range t.Rows {
			colCount = max(colCount, len(r.Cells))
		}
		ranges = append(ranges, databaseRange{
			Name:                 fmt.Sprintf("__Anonymous_Sheet_DB__%d", i),
			TargetRangeAddress:   usedRangeAddress(t.Name, rowCount, colCount),
			DisplayFilterButtons: "true",
		})
	}
	if len(ranges) == 0 {
		spreadsheet.DatabaseRanges = nil
		return spreadsheet
	}
	spreadsheet.DatabaseRanges = &databaseRanges{Ranges: ranges}
	return spreadsheet
}

// usedRangeAddress returns the table:target-range-address covering rowCount
// rows and colCount columns of the sheet, starting at A1 — e.g.
// "Sheet1.A1:Sheet1.B3". Unlike named ranges, database ranges use unanchored
// (no dollar sign) cell references.
func usedRangeAddress(sheet string, rowCount, colCount int) string {
	return fmt.Sprintf("%s.A1:%s.%s%d", sheet, sheet, columnToLetters(colCount), rowCount)
}

// customStyleKey identifies a distinct generated cell style, so cells
// sharing the same CellStyle and base data style reuse one style definition.
type customStyleKey struct {
	CellStyle
	dataStyleName string
}

// dataStyleNameFor maps a preset style name assigned by createCell (e.g.
// FLOAT_STYLE) to the number-format style it references, so custom styles
// keep the same numeric formatting. Value types without a preset style
// (string, time, percentage, formula) return "".
func dataStyleNameFor(presetStyleName string) string {
	switch presetStyleName {
	case "FLOAT_STYLE":
		return "FLOAT_DATA_STYLE"
	case "DATE_STYLE":
		return "DATE_DATA_STYLE"
	case "TIME_STYLE":
		return "TIME_DATA_STYLE"
	case "PERCENTAGE_STYLE":
		return "PERCENTAGE_DATA_STYLE"
	case "EUR_STYLE":
		return "EUR_DATA_STYLE"
	case "USD_STYLE":
		return "USD_DATA_STYLE"
	case "GBP_STYLE":
		return "GBP_DATA_STYLE"
	default:
		return ""
	}
}

func buildCustomCellStyle(name, dataStyleName string, style CellStyle) cellStyle {
	cs := cellStyle{
		Name:            name,
		Family:          "table-cell",
		ParentStyleName: "Default",
		DataStyleName:   dataStyleName,
	}
	if style.BackgroundColor != "" || style.Border != "" {
		cs.TableCellProperties = &tableCellProperties{
			BackgroundColor: style.BackgroundColor,
			Border:          style.Border,
		}
	}
	if style.FontColor != "" || style.Bold || style.Italic {
		tp := &textProperties{Color: style.FontColor}
		if style.Bold {
			tp.FontWeight = "bold"
		}
		if style.Italic {
			tp.FontStyle = "italic"
		}
		cs.TextProperties = tp
	}
	return cs
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
	pageStyles, master := createPageStyles()
	fods := flatOds{
		XMLNSOffice:    "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:     "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:      "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:     "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:        "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:    "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSMeta:      "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		XMLNSOf:        "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		XMLNSSvg:       "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		OfficeVersion:  odfVersion,
		OfficeMimetype: "application/vnd.oasis.opendocument.spreadsheet",
		Meta:           officeMeta{Generator: generator},
		Styles:         createCommonStyles(),
		AutomaticStyles: automaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createAutomaticStyles(spreadsheet.customStyles),
			PageLayout:   &pageStyles.PageLayout,
		},
		MasterStyles: master,
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
		XMLNSOf:       "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		OfficeVersion: odfVersion,
		AutomaticStyles: automaticStyles{
			NumberStyles: createNumberStyles(),
			Styles:       createAutomaticStyles(spreadsheet.customStyles),
		},
		Body: documentBody{
			Spreadsheet: spreadsheet,
		},
	}

	pageStyles, master := createPageStyles()
	stylesXml := documentStyles{
		XMLNSOffice:     "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		XMLNSTable:      "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		XMLNSText:       "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		XMLNSStyle:      "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		XMLNSFo:         "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		XMLNSNumber:     "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		XMLNSSvg:        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		OfficeVersion:   odfVersion,
		Styles:          createCommonStyles(),
		AutomaticStyles: pageStyles,
		MasterStyles:    master,
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
		writer, err := zipWriter.CreateHeader(&zip.FileHeader{
			Name:   part.name,
			Method: zip.Deflate,
			// Without an explicit timestamp the entries carry a zero time,
			// which the MS-DOS date fields of a zip archive cannot express;
			// they end up dated 1979-12-31, and strict readers reject that.
			// The timestamp is fixed rather than time.Now() to keep the
			// output byte-for-byte reproducible.
			Modified: zipEntryTime,
		})
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
		style:     data.Style,
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
		cell.StyleName = "TIME_STYLE"
		cell.TimeValue, cell.err = timeString(data.Value)
	case "percentage":
		cell.StyleName = "PERCENTAGE_STYLE"
		cell.err = parseNumber(data.Value, data.ValueType)
		cell.Value = data.Value
	case "formula":
		cell.Formula = toOpenFormula(data.Value)
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
		// Time and percentage carry an explicit format for the same reason
		// the other types do: a value without a data style is left to the
		// consumer to format, and only LibreOffice fills the gap with the
		// locale's default. Others pick something arbitrary — a time shown
		// with a date format renders as a day in 1899 — or nothing at all.
		// Neither style fixes the language, so the decimal and time
		// separators still follow the locale; only the format does not.
		timeStyle{
			Name: "TIME_DATA_STYLE",
			Parts: []any{
				timeHours{Style: "long"},
				textElement{Content: ":"},
				timeMinutes{Style: "long"},
				textElement{Content: ":"},
				timeSeconds{Style: "long"},
			},
		},
		percentageStyle{
			Name: "PERCENTAGE_DATA_STYLE",
			Number: numberElement{
				DecimalPlaces:    "2",
				MinDecimalPlaces: "2",
				MinIntegerDigits: "1",
			},
			Text: textElement{Content: "%"},
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
	// Only the style referenced through style:map is volatile. Leaving the
	// flag on the style the cells refer to marks it as unused, and consumers
	// are free to discard it along with the currency formatting.
	style.Volatile = ""
	style.TextProperties = &textProperties{Color: "#ff0000"}
	style.Texts = []textElement{{"−"}}
	style.StyleMap = &styleMap{Condition: "value()>=0", ApplyStyleName: code + "_DATA_STYLE_POSITIVE"}
	return style
}

// Names of the common styles and of the page setup shared by all sheets.
const (
	defaultCellStyleName = "Default"
	pageLayoutName       = "PAGE_LAYOUT"
	masterPageName       = "Default"
	tableStyleName       = "TABLE_STYLE"
)

// createCommonStyles returns the office:styles content: the default cell
// style all generated styles inherit from. Cell styles referring to a
// "Default" parent that is defined nowhere are silently dropped by some
// consumers, taking the number formats attached to them with it.
func createCommonStyles() officeStyles {
	return officeStyles{
		DefaultStyle: defaultCell{Family: "table-cell"},
		Styles: []cellStyle{
			{Name: defaultCellStyleName, Family: "table-cell"},
		},
	}
}

// createPageStyles returns the A4 portrait page layout and the master page
// referring to it.
func createPageStyles() (pageAutomaticStyle, masterStyles) {
	layout := pageAutomaticStyle{
		PageLayout: pageLayout{
			Name: pageLayoutName,
			Properties: pageLayoutProperties{
				PageWidth:        "21.0cm",
				PageHeight:       "29.7cm",
				PrintOrientation: "portrait",
				MarginTop:        "2cm",
				MarginBottom:     "2cm",
				MarginLeft:       "2cm",
				MarginRight:      "2cm",
			},
		},
	}
	master := masterStyles{
		MasterPage: masterPage{Name: masterPageName, PageLayoutName: pageLayoutName},
	}
	return layout, master
}

// createAutomaticStyles returns the automatic style:style definitions of a
// document: the preset cell styles, the styles generated for cells created
// with [MakeStyledCell], and the table style binding every sheet to the
// master page.
func createAutomaticStyles(customStyles []cellStyle) []any {
	var styles []any
	for _, style := range append(createStyles(), customStyles...) {
		styles = append(styles, style)
	}
	return append(styles, tableStyle{
		Name:           tableStyleName,
		Family:         "table",
		MasterPageName: masterPageName,
		Properties:     tableProperties{Display: "true"},
	})
}

func createStyles() []cellStyle {
	return []cellStyle{
		{Name: "FLOAT_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "FLOAT_DATA_STYLE"},
		{Name: "DATE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "DATE_DATA_STYLE"},
		{Name: "TIME_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "TIME_DATA_STYLE"},
		{Name: "PERCENTAGE_STYLE", Family: "table-cell", ParentStyleName: "Default", DataStyleName: "PERCENTAGE_DATA_STYLE"},
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
	style     *CellStyle
	err       error
}

// Spreadsheet is a collection of tables ready for serialization. Create
// spreadsheets with [MakeSpreadsheet]; the fields are exported only for XML
// marshaling.
type Spreadsheet struct {
	XMLName          xml.Name         `xml:"office:spreadsheet"`
	Tables           []table          `xml:"table:table"`
	NamedExpressions namedExpressions `xml:"table:named-expressions"`

	// DatabaseRanges holds AutoFilter definitions enabled via
	// [EnableAutoFilter]. The ODF schema requires table:database-ranges to
	// follow table:named-expressions, so this field is declared after it.
	DatabaseRanges *databaseRanges `xml:"table:database-ranges,omitempty"`

	// customStyles holds the cell styles generated for cells created with
	// [MakeStyledCell]. It is emitted into office:automatic-styles by
	// [MakeFlatOds] and [WriteOds].
	customStyles []cellStyle
}

// cellData is the raw input for a cell before validation.
type cellData struct {
	Value     string
	ValueType string
	Range     string
	Style     *CellStyle
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
	XMLName   xml.Name      `xml:"table:table"`
	Name      string        `xml:"table:name,attr"`
	StyleName string        `xml:"table:style-name,attr,omitempty"`
	Columns   []tableColumn `xml:"table:table-column"`
	Rows      []row         `xml:"table:table-row"`
}

// Field order matters throughout the document types: the ODF schema
// prescribes the order of the top-level elements, with office:styles before
// office:automatic-styles before office:master-styles.
type flatOds struct {
	XMLName         xml.Name        `xml:"office:document"`
	XMLNSOffice     string          `xml:"xmlns:office,attr"`
	XMLNSTable      string          `xml:"xmlns:table,attr"`
	XMLNSText       string          `xml:"xmlns:text,attr"`
	XMLNSStyle      string          `xml:"xmlns:style,attr"`
	XMLNSFo         string          `xml:"xmlns:fo,attr"`
	XMLNSNumber     string          `xml:"xmlns:number,attr"`
	XMLNSMeta       string          `xml:"xmlns:meta,attr"`
	XMLNSOf         string          `xml:"xmlns:of,attr"`
	XMLNSSvg        string          `xml:"xmlns:svg,attr"`
	OfficeVersion   string          `xml:"office:version,attr"`
	OfficeMimetype  string          `xml:"office:mimetype,attr"`
	Meta            officeMeta      `xml:"office:meta"`
	Styles          officeStyles    `xml:"office:styles"`
	AutomaticStyles automaticStyles `xml:"office:automatic-styles"`
	MasterStyles    masterStyles    `xml:"office:master-styles"`
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
	XMLNSOf         string          `xml:"xmlns:of,attr"`
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
	XMLName         xml.Name           `xml:"office:document-styles"`
	XMLNSOffice     string             `xml:"xmlns:office,attr"`
	XMLNSTable      string             `xml:"xmlns:table,attr"`
	XMLNSText       string             `xml:"xmlns:text,attr"`
	XMLNSStyle      string             `xml:"xmlns:style,attr"`
	XMLNSFo         string             `xml:"xmlns:fo,attr"`
	XMLNSNumber     string             `xml:"xmlns:number,attr"`
	XMLNSSvg        string             `xml:"xmlns:svg,attr"`
	OfficeVersion   string             `xml:"office:version,attr"`
	Styles          officeStyles       `xml:"office:styles"`
	AutomaticStyles pageAutomaticStyle `xml:"office:automatic-styles"`
	MasterStyles    masterStyles       `xml:"office:master-styles"`
}

type automaticStyles struct {
	XMLName      xml.Name `xml:"office:automatic-styles"`
	NumberStyles []any    `xml:"number:number-style"`
	// Styles holds the generated cell styles and the table style, which the
	// ODF schema both spells style:style.
	Styles []any `xml:"style:style"`

	// PageLayout is set only for flat documents, which have a single
	// office:automatic-styles for what the package splits over content.xml
	// and styles.xml.
	PageLayout *pageLayout `xml:"style:page-layout,omitempty"`
}

// officeStyles holds the common styles of a document. Every cell style this
// package generates inherits from the "Default" cell style defined here;
// consumers other than LibreOffice do not invent it when it is missing.
type officeStyles struct {
	XMLName      xml.Name    `xml:"office:styles"`
	DefaultStyle defaultCell `xml:"style:default-style"`
	Styles       []cellStyle `xml:"style:style"`
}

type defaultCell struct {
	XMLName xml.Name `xml:"style:default-style"`
	Family  string   `xml:"style:family,attr"`
}

// pageAutomaticStyle carries the page layout referenced by the master page.
type pageAutomaticStyle struct {
	XMLName    xml.Name   `xml:"office:automatic-styles"`
	PageLayout pageLayout `xml:"style:page-layout"`
}

type pageLayout struct {
	XMLName    xml.Name             `xml:"style:page-layout"`
	Name       string               `xml:"style:name,attr"`
	Properties pageLayoutProperties `xml:"style:page-layout-properties"`
}

type pageLayoutProperties struct {
	XMLName          xml.Name `xml:"style:page-layout-properties"`
	PageWidth        string   `xml:"fo:page-width,attr"`
	PageHeight       string   `xml:"fo:page-height,attr"`
	PrintOrientation string   `xml:"style:print-orientation,attr"`
	MarginTop        string   `xml:"fo:margin-top,attr"`
	MarginBottom     string   `xml:"fo:margin-bottom,attr"`
	MarginLeft       string   `xml:"fo:margin-left,attr"`
	MarginRight      string   `xml:"fo:margin-right,attr"`
}

// masterStyles holds the master page every sheet refers to through its table
// style. Spreadsheet applications take the page setup used for printing from
// it, and Excel expects one to be present.
type masterStyles struct {
	XMLName    xml.Name   `xml:"office:master-styles"`
	MasterPage masterPage `xml:"style:master-page"`
}

type masterPage struct {
	XMLName        xml.Name `xml:"style:master-page"`
	Name           string   `xml:"style:name,attr"`
	PageLayoutName string   `xml:"style:page-layout-name,attr"`
}

// tableStyle is the automatic style of a sheet, binding it to the master page.
type tableStyle struct {
	XMLName        xml.Name        `xml:"style:style"`
	Name           string          `xml:"style:name,attr"`
	Family         string          `xml:"style:family,attr"`
	MasterPageName string          `xml:"style:master-page-name,attr"`
	Properties     tableProperties `xml:"style:table-properties"`
}

type tableProperties struct {
	XMLName xml.Name `xml:"style:table-properties"`
	Display string   `xml:"table:display,attr"`
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
	Color      string `xml:"fo:color,attr,omitempty"`
	FontWeight string `xml:"fo:font-weight,attr,omitempty"`
	FontStyle  string `xml:"fo:font-style,attr,omitempty"`
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
	XMLName             xml.Name             `xml:"style:style"`
	Name                string               `xml:"style:name,attr"`
	Family              string               `xml:"style:family,attr"`
	ParentStyleName     string               `xml:"style:parent-style-name,attr,omitempty"`
	DataStyleName       string               `xml:"style:data-style-name,attr,omitempty"`
	TableCellProperties *tableCellProperties `xml:"style:table-cell-properties,omitempty"`
	TextProperties      *textProperties      `xml:"style:text-properties,omitempty"`
}

// tableCellProperties holds the visual cell properties generated for
// [CellStyle] (background color and border).
type tableCellProperties struct {
	BackgroundColor string `xml:"fo:background-color,attr,omitempty"`
	Border          string `xml:"fo:border,attr,omitempty"`
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

type timeStyle struct {
	XMLName xml.Name `xml:"number:time-style"`
	Name    string   `xml:"style:name,attr"`
	Parts   []any    `xml:"number:text"`
}

type timeHours struct {
	XMLName xml.Name `xml:"number:hours"`
	Style   string   `xml:"number:style,attr"`
}

type timeMinutes struct {
	XMLName xml.Name `xml:"number:minutes"`
	Style   string   `xml:"number:style,attr"`
}

type timeSeconds struct {
	XMLName xml.Name `xml:"number:seconds"`
	Style   string   `xml:"number:style,attr"`
}

// Field order matters: the ODF schema requires the number:number of a
// percentage style to precede the number:text carrying the percent sign.
type percentageStyle struct {
	XMLName xml.Name      `xml:"number:percentage-style"`
	Name    string        `xml:"style:name,attr"`
	Number  numberElement `xml:"number:number"`
	Text    textElement   `xml:"number:text"`
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

type databaseRanges struct {
	XMLName xml.Name        `xml:"table:database-ranges"`
	Ranges  []databaseRange `xml:"table:database-range"`
}

type databaseRange struct {
	XMLName              xml.Name `xml:"table:database-range"`
	Name                 string   `xml:"table:name,attr,omitempty"`
	TargetRangeAddress   string   `xml:"table:target-range-address,attr"`
	DisplayFilterButtons string   `xml:"table:display-filter-buttons,attr,omitempty"`
}

type namedExpressions struct {
	NamedRanges []namedRange `xml:"table:named-range"`
}

type namedRange struct {
	Name             string `xml:"table:name,attr"`
	BaseCellAddress  string `xml:"table:base-cell-address,attr"`
	CellRangeAddress string `xml:"table:cell-range-address,attr"`
}
