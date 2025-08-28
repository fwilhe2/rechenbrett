// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"encoding/xml"
	"strings"
)

type Cell struct {
	XMLName     xml.Name `xml:"table:table-cell"`
	Text        string   `xml:"text:p,omitempty"`
	ValueType   string   `xml:"office:value-type,attr,omitempty"`
	CalcExtType string   `xml:"calcext:value-type,attr,omitempty"`
	Value       string   `xml:"office:value,attr,omitempty"`
	DateValue   string   `xml:"office:date-value,attr,omitempty"`
	TimeValue   string   `xml:"office:time-value,attr,omitempty"`
	Currency    string   `xml:"office:currency,attr,omitempty"`
	StyleName   string   `xml:"table:style-name,attr,omitempty"`
	Formula     string   `xml:"table:formula,attr,omitempty"`
	Range       string   `xml:"-"`
}

type CellData struct {
	Value     string `json:"value"`
	ValueType string `json:"valueType"`
	Range     string `json:"range,omitempty"`
}



func createCell(cellData CellData) Cell {
	cell := Cell{
		ValueType: cellData.ValueType,
		Range:     cellData.Range,
	}

	switch cellData.ValueType {
	case "string":
		cell.Text = cellData.Value
		cell.CalcExtType = "string"
	case "float":
		cell.CalcExtType = "float"
		cell.StyleName = "FLOAT_STYLE"
		cell.Value = cellData.Value
	case "date":
		cell.CalcExtType = "date"
		cell.StyleName = "DATE_STYLE"
		cell.DateValue = dateString(cellData.Value)
	case "time":
		cell.CalcExtType = "time"
		cell.StyleName = "TIME_STYLE"
		cell.TimeValue = timeString(cellData.Value)
	case "percentage":
		cell.CalcExtType = "percentage"
		cell.StyleName = "PERCENTAGE_STYLE"
		cell.Value = cellData.Value
	case "formula":
		cell.Formula = cellData.Value
		cell.ValueType = ""
	default:
		if strings.HasPrefix(cellData.ValueType, "currency") {
			if strings.HasSuffix(strings.ToLower(cellData.ValueType), "usd") {
				cell.CalcExtType = "currency"
				cell.StyleName = "USD_STYLE"
				cell.Value = cellData.Value
				cell.Currency = "USD"
			} else if strings.HasSuffix(strings.ToLower(cellData.ValueType), "gbp") {
				cell.CalcExtType = "currency"
				cell.StyleName = "GBP_STYLE"
				cell.Value = cellData.Value
				cell.Currency = "GBP"
			} else {
				// Assuming Euro as the default, just because it is the default for me :shrug:
				cell.CalcExtType = "currency"
				cell.StyleName = "EUR_STYLE"
				cell.Value = cellData.Value
				cell.Currency = "EUR"
			}
		}
	}
	return cell
}