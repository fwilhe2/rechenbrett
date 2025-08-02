// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"os/exec"
	"strings"
	"testing"
)

func assert(t *testing.T, condition bool, message string) {
	if !condition {
		t.Error(message)
	}
}

func integrationTest(testName, format string, inputCells [][]Cell, expectedCsv map[string][][]string) error {
	lang := os.Getenv("LANG")
	spreadsheet := MakeSpreadsheet(inputCells)

	tempDir, err := os.MkdirTemp(".", fmt.Sprintf("integration-test-%s-%s-", testName, lang))
	if err != nil {
		panic(err)
	}

	if format == "ods" {
		buff := MakeOds(spreadsheet)

		archive, err := os.Create(fmt.Sprintf("%s/%s-%s.%s", tempDir, testName, lang, format))
		if err != nil {
			panic(err)
		}

		archive.Write(buff.Bytes())

		archive.Close()
	} else {
		actual := MakeFlatOds(spreadsheet)
		os.WriteFile(fmt.Sprintf("%s/%s-%s.%s", tempDir, testName, lang, format), []byte(actual), 0o644)
	}

	cmd := fmt.Sprintf("libreoffice --headless --convert-to csv:\"Text - txt - csv (StarCalc)\":\"44,34,76,1,,1031,true,true\" %s/%s-%s.%s --outdir %s", tempDir, testName, lang, format, tempDir)
	loCmd := exec.Command("bash", "-c", cmd)
	_, err = loCmd.Output()
	if err != nil {
		panic(err)
	}

	actualCsvBytes, _ := os.ReadFile(fmt.Sprintf("%s/%s-%s.csv", tempDir, testName, lang))

	actualCsv := string(actualCsvBytes)

	r := csv.NewReader(strings.NewReader(actualCsv))

	line := 0
	for {
		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Fatal(err)
		}

		fmt.Println(record)
		fmt.Println(expectedCsv)

		for i, v := range record {
			if v != expectedCsv[lang][line][i] {
				return fmt.Errorf("[%s %s] Failed test, value is: '%s', expected: '%s'", testName, lang, v, expectedCsv[lang][line][i])
			}
		}
	}
	return nil
}

func TestCommonDataTypes(t *testing.T) {
	givenThoseCells := [][]Cell{
		{
			MakeCell("ABBA", "string"),
			MakeCell("42.3324", "float"),
			MakeCell("-42.3324", "float"),
			MakeCell("2022-02-02", "date"),
			MakeCell("2.2.2022", "date"),
			MakeCell("19:03:00", "time"),
			MakeCell("2.22", "currency"),
			MakeCell("-2.22", "currency"),
			MakeCell("2.22", "currency-usd"),
			MakeCell("2.22", "currency-gbp"),
			MakeCell("2.22", "currency-eur"),
			MakeCell("0.4223", "percentage"),
		},
	}

	expectedThisCsv := make(map[string][][]string)
	expectedThisCsv["en_US.UTF-8"] = [][]string{
		{
			"ABBA",
			"42.33",
			"−42.33",
			"2022-02-02",
			"2022-02-02",
			"07:03:00 PM",
			"2.22€",
			"-2.22€",
			"2.22$",
			"2.22£",
			"2.22€",
			"42.23%",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"ABBA",
			"42,33",
			"−42,33",
			"2022-02-02",
			"2022-02-02",
			"19:03:00",
			"2.22€",
			"-2.22€",
			"2.22$",
			"2.22£",
			"2.22€",
			"42,23 %",
		},
	}

	err := integrationTest("common-data-types", "ods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))

	err = integrationTest("common-data-types", "fods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))
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
			"-2.00€",
			"2.20€",
			"-2.20€",
			"2.22€",
			"-2.22€",
		},
	}
	expectedThisCsv["de_DE.UTF-8"] = [][]string{
		{
			"2.00€",
			"-2.00€",
			"2.20€",
			"-2.20€",
			"2.22€",
			"-2.22€",
		},
	}

	err := integrationTest("currency-formatting", "ods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))

	err = integrationTest("currency-formatting", "fods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))
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

	err := integrationTest("formula", "ods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))

	err = integrationTest("formula", "fods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))
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

	err := integrationTest("ranges", "ods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))

	err = integrationTest("ranges", "fods", givenThoseCells, expectedThisCsv)
	assert(t, err == nil, fmt.Sprintf("err: %v\n", err))
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

	spreadsheet := MakeSpreadsheet(givenThoseCells)

	actual := MakeFlatOds(spreadsheet)

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
		assert(t, timeString(candidate) == expected, fmt.Sprintf("Expected %s to be parsed as %s", candidate, expected))
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
		assert(t, dateString((candidate)) == expected, fmt.Sprintf("Expected %s to be formatted as %s", candidate, expected))
	}
}
