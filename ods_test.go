package ods

import (
	"fmt"
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
}
