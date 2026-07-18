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

	lang := os.Getenv("LANG")
	expected, ok := expectedCsv[lang]
	if !ok {
		t.Skipf("no expectations for LANG=%q, set LANG to one of the supported locales", lang)
	}

	spreadsheet, err := MakeSpreadsheet(inputCells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
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
		"--convert-to", `csv:"Text - txt - csv (StarCalc)":"44,34,76,1,,1031,true,true"`,
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
			"07:03:00 PM",
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
			"42,23 %",
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

