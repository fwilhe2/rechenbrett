package ods

import (
	"fmt"
	"os"
	"os/exec"
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
			MakeCell("0.4223", "percentage"),
		},
	}

	spreadsheet := MakeSpreadsheet(cells)

	actual := MakeFlatOds(spreadsheet)

	fmt.Println(actual)

	testName := "commonDataTypes"

	os.Mkdir("output", 0777)
	os.WriteFile("output/"+testName+".fods", []byte(actual), 0644)

	cmd := fmt.Sprintf("libreoffice --headless --convert-to csv:\"Text - txt - csv (StarCalc)\":\"44,34,76,1,,1031,true,true\" output/%s.fods --outdir output", testName)
	loCmd := exec.Command("bash", "-c", cmd)
	_, err := loCmd.Output()
	if err != nil {
		panic(err)
	}

	csv, _ := os.ReadFile("output/" + testName + ".csv")

	fmt.Println(string(csv))
	if string(csv) != "\"ABBA\",\"42,33\",02.02.22 00:00,19:03:00,2.22,\"42,23 %\"" {
		// t.Fail()
		fmt.Println("todo: assert")
	}
}
