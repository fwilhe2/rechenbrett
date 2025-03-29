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

func integrationTest(testName string, format string, inputCells [][]Cell, expectedCsv [][]string) error {
	spreadsheet := MakeSpreadsheet(inputCells)

	actual := MakeFlatOds(spreadsheet)
	os.Mkdir("output", 0777)

	if format == "ods" {
		buff := MakeOds(spreadsheet)

		archive, err := os.Create(fmt.Sprintf("output/%s.%s", testName, format))
		if err != nil {
			panic(err)
		}

		archive.Write(buff.Bytes())

		archive.Close()
	} else {
		os.WriteFile(fmt.Sprintf("output/%s.%s", testName, format), []byte(actual), 0644)
	}

	cmd := fmt.Sprintf("libreoffice --headless --convert-to csv:\"Text - txt - csv (StarCalc)\":\"44,34,76,1,,1031,true,true\" output/%s.%s --outdir output", testName, format)
	loCmd := exec.Command("bash", "-c", cmd)
	_, err := loCmd.Output()
	if err != nil {
		panic(err)
	}

	actualCsvBytes, _ := os.ReadFile(fmt.Sprintf("output/%s.csv", testName))

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

		for i, v := range record {
			if v != expectedCsv[line][i] {
				return fmt.Errorf("[%s] Failed test, value is: '%s', expected: '%s'", testName, v, expectedCsv[line][i])
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
			MakeCell("2022-02-02", "date"),
			MakeCell("19:03:00", "time"),
			MakeCell("2.22", "currency"),
			MakeCell("-2.22", "currency"),
			MakeCell("0.4223", "percentage"),
		},
	}

	expectedThisCsv := [][]string{
		{
			"ABBA",
			"42,33",
			"2022-02-02",
			"19:03:00",
			"2.22€",
			"-2.22€",
			"42,23 %",
		},
	}

	err := integrationTest("common-data-types", "ods", givenThoseCells, expectedThisCsv)

	if err != nil {
		fmt.Printf("err: %v\n", err)
		t.Fail()
	}

	err = integrationTest("common-data-types", "fods", givenThoseCells, expectedThisCsv)

	if err != nil {
		fmt.Printf("err: %v\n", err)
		t.Fail()
	}
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

	expectedThisCsv := [][]string{
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

	if err != nil {
		fmt.Printf("err: %v\n", err)
		t.Fail()
	}

	err = integrationTest("formula", "fods", givenThoseCells, expectedThisCsv)

	if err != nil {
		fmt.Printf("err: %v\n", err)
		t.Fail()
	}
}
