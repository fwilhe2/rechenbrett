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

func integrationTest(testName string, inputCells [][]Cell, expectedCsv [][]string) error {
	spreadsheet := MakeSpreadsheet(inputCells)

	actual := MakeFlatOds(spreadsheet)
	os.Mkdir("output", 0777)

	os.WriteFile(fmt.Sprintf("output/%s.fods", testName), []byte(actual), 0644)

	cmd := fmt.Sprintf("libreoffice --headless --convert-to csv:\"Text - txt - csv (StarCalc)\":\"44,34,76,1,,1031,true,true\" output/%s.fods --outdir output", testName)
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

	err := integrationTest("common-data-types", givenThoseCells, expectedThisCsv)

	if err != nil {
		fmt.Printf("err: %v\n", err)
		t.Fail()
	}
}


func TestFoo(t *testing.T) {
	inputCells := [][]Cell{
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

	spreadsheet := MakeSpreadsheet(inputCells)


	buff := MakeOds(spreadsheet)

	archive, err := os.Create("output/test.ods")
	if err != nil {
		panic(err)
	}

	archive.Write(buff.Bytes())

	archive.Close()
}