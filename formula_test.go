// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import "testing"

func TestUnitToOpenFormula(t *testing.T) {
	cases := map[string]string{
		// Cell references become bracketed and sheet-qualified.
		"A1+B1":                    "of:=[.A1]+[.B1]",
		"SUM(A1:B1)":               "of:=SUM([.A1:.B1])",
		"(A1+B1)/2":                "of:=([.A1]+[.B1])/2",
		"AA10*2":                   "of:=[.AA10]*2",
		"$A$1+B$2":                 "of:=[.$A$1]+[.B$2]",
		"Sheet2.A1":                "of:=[Sheet2.A1]",
		"SUM(Sheet2.A1:Sheet2.B2)": "of:=SUM([Sheet2.A1:Sheet2.B2])",

		// Named ranges are referenced bare in OpenFormula, so they are left
		// as they are.
		"InputA+InputB":      "of:=InputA+InputB",
		"SUM(InputA:InputB)": "of:=SUM(InputA:InputB)",
		"AVERAGE(A1:B1)":     "of:=AVERAGE([.A1:.B1])",

		// Function arguments are separated by semicolons.
		"SUM(A1,B1)":      "of:=SUM([.A1];[.B1])",
		"IF(A1>0,A1,-A1)": "of:=IF([.A1]>0;[.A1];-[.A1])",

		// Whatever looks like a reference inside a string literal is not one.
		`CONCATENATE("A1, B1",A1)`: `of:=CONCATENATE("A1, B1";[.A1])`,

		// Numbers are not references, and their decimal point is not a sheet
		// separator.
		"A1*1.5": "of:=[.A1]*1.5",

		// A leading equals sign is accepted and dropped, an expression that
		// already carries a namespace prefix is passed through.
		"=A1+B1":         "of:=[.A1]+[.B1]",
		"of:=SUM([.A1])": "of:=SUM([.A1])",
	}

	for input, expected := range cases {
		t.Run(input, func(t *testing.T) {
			if actual := toOpenFormula(input); actual != expected {
				t.Errorf("toOpenFormula(%q) = %q, expected %q", input, actual, expected)
			}
		})
	}
}
