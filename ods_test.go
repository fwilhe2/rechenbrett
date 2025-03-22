package ods

import (
	"fmt"
	"testing"
)

func TestMakeSpreadsheetWithCommonDataTypes(t *testing.T) {
	var allCells [][]Cell
	var firstRowCells []Cell

	firstRowCells = append(firstRowCells, createCell(CellData{Value: "ABBA", ValueType: "string"}))
	firstRowCells = append(firstRowCells, createCell(CellData{Value: "42.3324", ValueType: "float"}))
	firstRowCells = append(firstRowCells, createCell(CellData{Value: "2022-02-02", ValueType: "date"}))
	firstRowCells = append(firstRowCells, createCell(CellData{Value: "19:03:00", ValueType: "time"}))
	firstRowCells = append(firstRowCells, createCell(CellData{Value: "2.22", ValueType: "currency"}))
	firstRowCells = append(firstRowCells, createCell(CellData{Value: "0.4223", ValueType: "percentage"}))

	allCells = append(allCells, firstRowCells)
	spreadsheet := MakeSpreadsheet(allCells)

	actual := MakeFlatOds(spreadsheet)

	fmt.Println(actual)
}
