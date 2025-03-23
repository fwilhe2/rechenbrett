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

func TestMakeSpreadsheetWithCommonDataTypes(t *testing.T) {

	cells := [][]Cell{
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

	spreadsheet := MakeSpreadsheet(cells)

	actual := MakeFlatOds(spreadsheet)

	fmt.Println(actual)

	testName := "commonDataTypes"

	os.Mkdir("output", 0777)

	os.WriteFile(fmt.Sprintf("output/%s.fods", testName), []byte(actual), 0644)

	cmd := fmt.Sprintf("libreoffice --headless --convert-to csv:\"Text - txt - csv (StarCalc)\":\"44,34,76,1,,1031,true,true\" output/%s.fods --outdir output", testName)
	loCmd := exec.Command("bash", "-c", cmd)
	_, err := loCmd.Output()
	if err != nil {
		panic(err)
	}

	actualCsv, _ := os.ReadFile(fmt.Sprintf("output/%s.csv", testName))

	fmt.Println(string(actualCsv))

	r := csv.NewReader(strings.NewReader(string(actualCsv)))

	expected := []string{"ABBA", "42,33", "2022-02-02", "19:03:00", "2.22€", "-2.22€", "42,23 %"}

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
			if (v != expected[i]) {
				fmt.Printf("Failed compare, is: %s, expected: %s", v, expected[i])
				t.Fail()
			}
		}
	}


}
