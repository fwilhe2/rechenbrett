// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"testing"
	"encoding/xml"
)

func TestCreateCell(t *testing.T) {
	expected := Cell{
		Text: "foo",
		ValueType: "string",
		CalcExtType: "string",
	}
	actual := createCell(CellData{
		Value: "foo",
		ValueType: "string",
	})
	if actual != expected {
		t.Fail()
	}
}

func TestXxx(t *testing.T) {
	expected := "<table:table-cell office:value-type=\"string\" calcext:value-type=\"string\"><text:p>foo</text:p></table:table-cell>"
	xx := Cell{
		Text: "foo",
		ValueType: "string",
		CalcExtType: "string",
	}
	actualBytes, err := xml.Marshal(xx)
	if err != nil {
		t.Fail()
	}
	actual := string(actualBytes)
	if actual != expected {
		t.Fail()
	}
}