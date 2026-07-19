// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"encoding/csv"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"os/exec"
	"regexp"
	"strings"
	"testing"
)

func assert(t *testing.T, condition bool, message string) {
	t.Helper()
	if !condition {
		t.Error(message)
	}
}

// integrationTest renders the cells to a file, has LibreOffice convert it to
// CSV, and compares the result with the locale-specific expectation.
//
// The temp dirs are deliberately created inside the repository and not
// cleaned up: the odfvalidator step in CI scans them and validates the
// generated ods files. Do not replace them with t.TempDir().
//
// These tests must not run in parallel: LibreOffice instances share a single
// user profile.
func integrationTest(t *testing.T, testName, format string, inputCells [][]Cell, expectedCsv map[string][][]string) {
	t.Helper()

	spreadsheet, err := MakeSpreadsheet(inputCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}
	renderAndCompare(t, testName, format, spreadsheet, expectedCsv)
}

// tableTest is like integrationTest but builds the document with MakeTable
// (styled header, banded rows, AutoFilter, and a SUBTOTAL sum in the last
// column's totals row), rendering it through LibreOffice for comparison.
func tableTest(t *testing.T, testName, format string, inputCells [][]Cell, expectedCsv map[string][][]string) {
	t.Helper()

	cols := 0
	for _, r := range inputCells {
		if len(r) > cols {
			cols = len(r)
		}
	}
	totals := make([]Total, cols)
	if cols > 0 {
		totals[cols-1] = Total{TotalSum}
	}

	spreadsheet, err := MakeTable(inputCells, TableOptions{
		Header:     true,
		AutoFilter: true,
		BandedRows: true,
		Totals:     totals,
	})
	if err != nil {
		t.Fatalf("MakeTable: %v", err)
	}
	renderAndCompare(t, testName, format, spreadsheet, expectedCsv)
}

// renderAndCompare writes spreadsheet to a file, has LibreOffice convert it to
// CSV, and compares the result with the locale-specific expectation.
func renderAndCompare(t *testing.T, testName, format string, spreadsheet Spreadsheet, expectedCsv map[string][][]string) {
	t.Helper()

	lang := os.Getenv("LANG")
	expected, ok := expectedCsv[lang]
	if !ok {
		t.Skipf("no expectations for LANG=%q, set LANG to one of the supported locales", lang)
	}

	tempDir, err := os.MkdirTemp(".", fmt.Sprintf("_it-%s-%s-%s-", testName, format, lang))
	if err != nil {
		t.Fatal(err)
	}

	filename := fmt.Sprintf("%s/%s-%s.%s", tempDir, testName, lang, format)
	if format == "ods" {
		buff, err := MakeOds(spreadsheet)
		if err != nil {
			t.Fatalf("MakeOds: %v", err)
		}
		if err := os.WriteFile(filename, buff.Bytes(), 0o644); err != nil {
			t.Fatal(err)
		}
	} else {
		actual, err := MakeFlatOds(spreadsheet)
		if err != nil {
			t.Fatalf("MakeFlatOds: %v", err)
		}
		if err := os.WriteFile(filename, []byte(actual), 0o644); err != nil {
			t.Fatal(err)
		}
	}

	loCmd := exec.Command("libreoffice", "--headless",
		"--convert-to", `csv:Text - txt - csv (StarCalc):44,34,76,1,,1031,true,true`,
		filename, "--outdir", tempDir)
	output, err := loCmd.CombinedOutput()
	if err != nil {
		t.Fatalf("[%s %s] libreoffice conversion failed: %v\n%s", testName, lang, err, output)
	}

	actualCsvBytes, err := os.ReadFile(fmt.Sprintf("%s/%s-%s.csv", tempDir, testName, lang))
	if err != nil {
		t.Fatal(err)
	}

	r := csv.NewReader(strings.NewReader(string(actualCsvBytes)))

	line := 0
	for {
		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatal(err)
		}

		for i, v := range record {
			if v != expected[line][i] {
				t.Errorf("[%s %s] line %d, column %d: got %q, expected %q", testName, lang, line+1, i+1, v, expected[line][i])
			}
		}
		line++
	}
}

func TestCommonDataTypes(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeCell("ABBA", "string"),
			MakeCell("42", "float"),
			MakeCell("-42", "float"),
			MakeCell("42.3324", "float"),
			MakeCell("-42.3324", "float"),
			MakeCell("2022-02-02", "date"),
			MakeCell("2.2.2022", "date"),
			MakeCell("19:03:00", "time"),
			MakeCell("2.22", "currency"),
			MakeCell("-2.22", "currency"),
			MakeCell("2.22", "currency-usd"),
			MakeCell("-2.22", "currency-usd"),
			MakeCell("2.22", "currency-gbp"),
			MakeCell("-2.22", "currency-gbp"),
			MakeCell("2.22", "currency-eur"),
			MakeCell("-2.22", "currency-eur"),
			MakeCell("0.4223", "percentage"),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{
			"ABBA",
			"42.00",
			"−42.00",
			"42.33",
			"−42.33",
			"2022-02-02",
			"2022-02-02",
			// Time and percentage carry an explicit data style, so the
			// format no longer follows the locale — only the separators do.
			"19:03:00",
			"2.22€",
			"−2.22€",
			"2.22$",
			"−2.22$",
			"2.22£",
			"−2.22£",
			"2.22€",
			"−2.22€",
			"42.23%",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"ABBA",
			"42,00",
			"−42,00",
			"42,33",
			"−42,33",
			"2022-02-02",
			"2022-02-02",
			"19:03:00",
			"2.22€",
			"−2.22€",
			"2.22$",
			"−2.22$",
			"2.22£",
			"−2.22£",
			"2.22€",
			"−2.22€",
			"42,23%",
		},
	}

	integrationTest(t, "common-data-types", "ods", givenThoseCells, expectedThisCsv)
	integrationTest(t, "common-data-types", "fods", givenThoseCells, expectedThisCsv)
}

func TestCurrencyFormatting(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeCell("2", "currency"),
			MakeCell("-2", "currency"),
			MakeCell("2.2", "currency"),
			MakeCell("-2.2", "currency"),
			MakeCell("2.22", "currency"),
			MakeCell("-2.22", "currency"),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{
			"2.00€",
			"−2.00€",
			"2.20€",
			"−2.20€",
			"2.22€",
			"−2.22€",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"2.00€",
			"−2.00€",
			"2.20€",
			"−2.20€",
			"2.22€",
			"−2.22€",
		},
	}

	integrationTest(t, "currency-formatting", "ods", givenThoseCells, expectedThisCsv)
	integrationTest(t, "currency-formatting", "fods", givenThoseCells, expectedThisCsv)
}

func TestFormula(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeCell("42.3324", "float"),
			MakeCell("23", "float"),
			MakeCell("A1+B1", "formula"),
			MakeCell("SUM(A1:B1)", "formula"),
			MakeCell("(A1+B1)/2", "formula"),
			MakeCell("AVERAGE(A1:B1)", "formula"),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{
			"42.33",
			"23.00",
			"65.3324",
			"65.3324",
			"32.6662",
			"32.6662",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"42,33",
			"23,00",
			"65,3324",
			"65,3324",
			"32,6662",
			"32,6662",
		},
	}

	integrationTest(t, "formula", "ods", givenThoseCells, expectedThisCsv)
	integrationTest(t, "formula", "fods", givenThoseCells, expectedThisCsv)
}

func TestRanges(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeRangeCell("42.3324", "float", "InputA"),
			MakeRangeCell("23", "float", "InputB"),
			MakeCell("InputA+InputB", "formula"),
			MakeCell("SUM(InputA:InputB)", "formula"),
			MakeCell("(InputA+InputB)/2", "formula"),
			MakeCell("AVERAGE(InputA:InputB)", "formula"),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{
			"42.33",
			"23.00",
			"65.3324",
			"65.3324",
			"32.6662",
			"32.6662",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"42,33",
			"23,00",
			"65,3324",
			"65,3324",
			"32,6662",
			"32,6662",
		},
	}

	integrationTest(t, "ranges", "ods", givenThoseCells, expectedThisCsv)
	integrationTest(t, "ranges", "fods", givenThoseCells, expectedThisCsv)
}

func TestUnitRanges(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeRangeCell("42.3324", "float", "InputA"),
			MakeRangeCell("23", "float", "InputB"),
		},
		{},
		{
			MakeRangeCell("42.3324", "float", "InputC"),
			MakeRangeCell("23", "float", "InputD"),
		},
		{},
		{
			MakeRangeCell("42.3324", "float", "InputE"),
			MakeRangeCell("23", "float", "InputF"),
		},
	}

	spreadsheet, err := MakeSpreadsheet(givenThoseCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, strings.Contains(actual, "<table:named-range table:name=\"InputA\" table:base-cell-address=\"$Sheet1.$A$1\" table:cell-range-address=\"$Sheet1.$A$1\"></table:named-range>"), "Expected input A in spreadsheet")
	assert(t, strings.Contains(actual, "<table:named-range table:name=\"InputF\" table:base-cell-address=\"$Sheet1.$B$5\" table:cell-range-address=\"$Sheet1.$B$5\"></table:named-range>"), "Expected input F in spreadsheet")
}

func TestUnitAutoFilter(t *testing.T) {
	givenThoseCells := [][]Cell{
		{MakeCell("Name", "string"), MakeCell("Age", "string")},
		{MakeCell("Alice", "string"), MakeCell("30", "float")},
		{MakeCell("Bob", "string"), MakeCell("25", "float")},
	}

	spreadsheet, err := MakeSpreadsheet(givenThoseCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}
	spreadsheet = EnableAutoFilter(spreadsheet)

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, strings.Contains(actual, "<table:database-ranges>"), "expected a database-ranges element")
	assert(t, strings.Contains(actual, `table:target-range-address="Sheet1.A1:Sheet1.B3"`), "expected the filter to cover the used range Sheet1.A1:Sheet1.B3")
	assert(t, strings.Contains(actual, `table:display-filter-buttons="true"`), "expected filter buttons to be enabled")
}

func TestUnitNoAutoFilterByDefault(t *testing.T) {
	spreadsheet, err := MakeSpreadsheet([][]Cell{{MakeCell("Name", "string")}})
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, !strings.Contains(actual, "database-range"), "expected no database-range without EnableAutoFilter")
}

func TestUnitTable(t *testing.T) {
	cells := [][]Cell{
		{MakeCell("Product", "string"), MakeCell("Price", "string")},
		{MakeCell("Pen", "string"), MakeCell("1.49", "float")},
		{MakeCell("Desk", "string"), MakeCell("189.00", "float")},
	}

	spreadsheet, err := MakeTable(cells, TableOptions{
		Name:       "Products",
		Header:     true,
		AutoFilter: true,
		BandedRows: true,
		Totals:     []Total{{TotalNone}, {TotalSum}},
	})
	if err != nil {
		t.Fatalf("MakeTable: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	// AutoFilter covers header + body (rows 1-3) but not the totals row (row 4).
	assert(t, strings.Contains(actual, `table:name="Products"`), "expected the database range named after the table")
	assert(t, strings.Contains(actual, `table:target-range-address="Sheet1.A1:Sheet1.B3"`), "expected the filter range to exclude the totals row")
	assert(t, strings.Contains(actual, `table:display-filter-buttons="true"`), "expected filter buttons")

	// Header row is styled bold with the blue theme's header fill.
	assert(t, strings.Contains(actual, `fo:background-color="#00599d"`), "expected the blue header fill")
	assert(t, strings.Contains(actual, `fo:font-weight="bold"`), "expected a bold header")
	// Banded body: the second body row (row 3) is filled.
	assert(t, strings.Contains(actual, `fo:background-color="#dddddd"`), "expected banded-row fill")
	// Totals row: a SUBTOTAL over the body of column B, with the totals fill.
	assert(t, strings.Contains(actual, `table:formula="of:=SUBTOTAL(9;[.B2:.B3])"`), "expected a SUBTOTAL sum over the body of column B")
	assert(t, strings.Contains(actual, `fo:background-color="#adc5e7"`), "expected the totals-row fill")
}

func TestUnitTablePlainByDefault(t *testing.T) {
	// The zero-value TableOptions produces a plain table: no header styling,
	// no filter, no banding, no totals.
	spreadsheet, err := MakeTable([][]Cell{
		{MakeCell("a", "string"), MakeCell("b", "string")},
		{MakeCell("1", "float"), MakeCell("2", "float")},
	}, TableOptions{})
	if err != nil {
		t.Fatalf("MakeTable: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, !strings.Contains(actual, "database-range"), "expected no AutoFilter for zero-value options")
	assert(t, !strings.Contains(actual, "CUSTOM_STYLE_"), "expected no generated styles for a plain table")
	assert(t, !strings.Contains(actual, "SUBTOTAL"), "expected no totals row for a plain table")
}

func TestUnitTableDoesNotMutateInput(t *testing.T) {
	cells := [][]Cell{
		{MakeCell("Product", "string"), MakeCell("Price", "string")},
		{MakeCell("Pen", "string"), MakeCell("1.49", "float")},
	}

	if _, err := MakeTable(cells, TableOptions{Header: true, BandedRows: true}); err != nil {
		t.Fatalf("MakeTable: %v", err)
	}

	for i, r := range cells {
		for j, c := range r {
			if c.style != nil {
				t.Errorf("MakeTable mutated the caller's cell at row %d, column %d", i+1, j+1)
			}
		}
	}
}

func TestUnitTableReportsInvalidCells(t *testing.T) {
	_, err := MakeTable([][]Cell{
		{MakeCell("x", "float"), MakeCell("y", "date")},
	}, TableOptions{Header: true})
	if err == nil || !strings.Contains(err.Error(), "float") || !strings.Contains(err.Error(), "date") {
		t.Errorf("expected both cell errors to be reported, got: %v", err)
	}
}

func TestTable(t *testing.T) {
	givenThoseCells := [][]Cell{
		{MakeCell("Product", "string"), MakeCell("Price", "string")},
		{MakeCell("Pen", "string"), MakeCell("1.49", "float")},
		{MakeCell("Desk", "string"), MakeCell("189.00", "float")},
	}

	// Rendered to CSV, the styled table looks like plain data plus the totals
	// row, which SUBTOTAL fills in with the sum of the body (1.49 + 189.00).
	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{"Product", "Price"},
		{"Pen", "1.49"},
		{"Desk", "189.00"},
		{"", "190.49"},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{"Product", "Price"},
		{"Pen", "1,49"},
		{"Desk", "189,00"},
		{"", "190,49"},
	}

	tableTest(t, "table", "ods", givenThoseCells, expectedThisCsv)
	tableTest(t, "table", "fods", givenThoseCells, expectedThisCsv)
}

func TestUnitTimeParse(t *testing.T) {
	expected := "PT19H03M00S"

	testTimes := []string{
		"19:03:00",
		"19:03",
	}

	for _, candidate := range testTimes {
		actual, err := timeString(candidate)
		if err != nil {
			t.Errorf("timeString(%q): %v", candidate, err)
		}
		assert(t, actual == expected, fmt.Sprintf("Expected %s to be parsed as %s, got %s", candidate, expected, actual))
	}
}

func TestUnitDateParse(t *testing.T) {
	expected := "1903-10-01"

	testDates := []string{
		"01.10.1903",
		"1.10.1903",
		"10/01/1903",
		"10/1/1903",
		"1903-10-01",
	}

	for _, candidate := range testDates {
		actual, err := dateString(candidate)
		if err != nil {
			t.Errorf("dateString(%q): %v", candidate, err)
		}
		assert(t, actual == expected, fmt.Sprintf("Expected %s to be formatted as %s, got %s", candidate, expected, actual))
	}
}

func TestUnitInvalidInput(t *testing.T) {
	invalidCells := map[string]Cell{
		"unknown value type":  MakeCell("42", "number"),
		"unknown currency":    MakeCell("42", "currency-chf"),
		"malformed date":      MakeCell("02.02.22", "date"),
		"malformed time":      MakeCell("25 o'clock", "time"),
		"non-numeric float":   MakeCell("fourtytwo", "float"),
		"non-numeric percent": MakeCell("42 %", "percentage"),
	}

	for name, cell := range invalidCells {
		t.Run(name, func(t *testing.T) {
			_, err := MakeSpreadsheet([][]Cell{{cell}})
			if err == nil {
				t.Errorf("expected an error for %s", name)
			}
		})
	}

	t.Run("duplicate range name", func(t *testing.T) {
		_, err := MakeSpreadsheet([][]Cell{{
			MakeRangeCell("1", "float", "Input"),
			MakeRangeCell("2", "float", "Input"),
		}})
		if err == nil || !strings.Contains(err.Error(), "duplicate range name") {
			t.Errorf("expected duplicate range name error, got: %v", err)
		}
	})

	t.Run("all errors are reported", func(t *testing.T) {
		_, err := MakeSpreadsheet([][]Cell{{
			MakeCell("x", "float"),
			MakeCell("y", "date"),
		}})
		if err == nil || !strings.Contains(err.Error(), "float") || !strings.Contains(err.Error(), "date") {
			t.Errorf("expected both cell errors to be reported, got: %v", err)
		}
	})
}

func TestStyledCell(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeStyledCell("Navy", "string", CellStyle{BackgroundColor: ColorNavy}),
			MakeStyledCell("Bold", "string", CellStyle{Bold: true}),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{{"Navy", "Bold"}}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{{"Navy", "Bold"}}

	integrationTest(t, "styled-cell", "ods", givenThoseCells, expectedThisCsv)
	integrationTest(t, "styled-cell", "fods", givenThoseCells, expectedThisCsv)
}

func TestUnitColorPalette(t *testing.T) {
	palette := map[string]string{
		"ColorNavy":    ColorNavy,
		"ColorBlue":    ColorBlue,
		"ColorAqua":    ColorAqua,
		"ColorTeal":    ColorTeal,
		"ColorPurple":  ColorPurple,
		"ColorFuchsia": ColorFuchsia,
		"ColorMaroon":  ColorMaroon,
		"ColorRed":     ColorRed,
		"ColorOrange":  ColorOrange,
		"ColorYellow":  ColorYellow,
		"ColorOlive":   ColorOlive,
		"ColorGreen":   ColorGreen,
		"ColorLime":    ColorLime,
		"ColorBlack":   ColorBlack,
		"ColorGray":    ColorGray,
		"ColorSilver":  ColorSilver,
		"ColorWhite":   ColorWhite,
	}

	hexColor := regexp.MustCompile(`^#[0-9a-f]{6}$`)
	seen := map[string]string{}
	for name, value := range palette {
		assert(t, hexColor.MatchString(value), fmt.Sprintf("%s = %q is not a lowercase #rrggbb hex color", name, value))
		if other, exists := seen[value]; exists {
			t.Errorf("%s and %s both have value %q, expected the palette to have distinct colors", name, other, value)
		}
		seen[value] = name
	}
}

func TestUnitStyledCellGeneratesStyle(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeStyledCell("Navy", "string", CellStyle{BackgroundColor: "#001f3f"}),
			MakeStyledCell("42", "float", CellStyle{Bold: true, FontColor: "#ffffff", Border: "0.5pt solid #000000"}),
		},
	}

	spreadsheet, err := MakeSpreadsheet(givenThoseCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, strings.Contains(actual, `table:style-name="CUSTOM_STYLE_1"`), "expected first styled cell to reference CUSTOM_STYLE_1")
	assert(t, strings.Contains(actual, `table:style-name="CUSTOM_STYLE_2"`), "expected second styled cell to reference CUSTOM_STYLE_2")
	assert(t, strings.Contains(actual, `style:name="CUSTOM_STYLE_1" style:family="table-cell" style:parent-style-name="Default"`), "expected CUSTOM_STYLE_1 style definition")
	assert(t, strings.Contains(actual, `fo:background-color="#001f3f"`), "expected background color in generated style")
	assert(t, strings.Contains(actual, `style:name="CUSTOM_STYLE_2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="FLOAT_DATA_STYLE"`), "expected the float cell's custom style to keep the float number format")
	assert(t, strings.Contains(actual, `fo:font-weight="bold"`), "expected bold font weight in generated style")
	assert(t, strings.Contains(actual, `fo:border="0.5pt solid #000000"`), "expected border in generated style")
}

func TestUnitStyledCellDeduplicates(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeStyledCell("A", "string", CellStyle{BackgroundColor: "#ff0000"}),
			MakeStyledCell("B", "string", CellStyle{BackgroundColor: "#ff0000"}),
		},
	}

	spreadsheet, err := MakeSpreadsheet(givenThoseCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}

	actual, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}

	assert(t, strings.Count(actual, `style:name="CUSTOM_STYLE_1"`) == 1, "expected only one style definition for two identically styled cells")
	assert(t, strings.Count(actual, `table:style-name="CUSTOM_STYLE_1"`) == 2, "expected both cells to reference the same generated style")
}

func TestUnitCell(t *testing.T) {
	givenThisCell := MakeCell("2.33", "float")
	expectThisXml := `<table:table-cell office:value-type="float" office:value="2.33" table:style-name="FLOAT_STYLE"></table:table-cell>`

	actualBytes, err := xml.Marshal(givenThisCell)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	actual := string(actualBytes)

	assert(t, actual == expectThisXml, fmt.Sprintf("Expected:\n%s\nGot:\n%s\n", expectThisXml, actual))
}

func TestUnitRow(t *testing.T) {
	cell1 := Cell{Value: "2"}
	cell2 := Cell{Value: "a"}
	givenThisRow := row{Cells: []Cell{cell1, cell2}}
	expectThisXml := `<table:table-row><table:table-cell office:value="2"></table:table-cell><table:table-cell office:value="a"></table:table-cell></table:table-row>`

	actualBytes, err := xml.Marshal(givenThisRow)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	actual := string(actualBytes)

	assert(t, actual == expectThisXml, fmt.Sprintf("Expected:\n%s\nGot:\n%s\n", expectThisXml, actual))
}
