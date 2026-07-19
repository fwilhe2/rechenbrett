// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
)

const schemaFile = "OpenDocument-v1.4-schema.rng"

// validateAgainstSchema validates one XML document against the ODF 1.4
// RELAX NG schema using jing. xmllint is not suitable: libxml2 cannot
// process the interleave patterns of the ODF schemas.
func validateAgainstSchema(t *testing.T, name, xmlContent string) {
	t.Helper()

	jing, err := exec.LookPath("jing")
	if err != nil {
		t.Skip("jing is not installed, skipping RELAX NG schema validation")
	}
	if _, err := os.Stat(schemaFile); err != nil {
		t.Skipf("schema %s not found, skipping RELAX NG schema validation", schemaFile)
	}

	path := filepath.Join(t.TempDir(), name)
	if err := os.WriteFile(path, []byte(xmlContent), 0o644); err != nil {
		t.Fatal(err)
	}

	// -i disables RELAX NG DTD-compatibility ID/IDREF checking, which the
	// published ODF schemas do not conform to.
	out, err := exec.Command(jing, "-i", schemaFile, path).CombinedOutput()
	if err != nil {
		t.Errorf("%s does not validate against the ODF 1.4 schema:\n%s", name, out)
	}
}

// schemaTestCases exercises every cell type the library can produce, plus
// notable variants (negative amounts, alternate date/time formats, named
// ranges), individually and combined. Each case is validated on its own so
// a schema violation can be traced back to the specific type that caused
// it, instead of being masked by unrelated cells in a shared spreadsheet.
var schemaTestCases = map[string][][]Cell{
	"string":                 {{MakeCell("ABBA", "string")}},
	"float":                  {{MakeCell("42.3324", "float")}},
	"float negative":         {{MakeCell("-42.3324", "float")}},
	"date iso":               {{MakeCell("2022-02-02", "date")}},
	"date german":            {{MakeCell("2.2.2022", "date")}},
	"date us":                {{MakeCell("2/2/2022", "date")}},
	"time hh:mm":             {{MakeCell("19:03", "time")}},
	"time hh:mm:ss":          {{MakeCell("19:03:00", "time")}},
	"percentage":             {{MakeCell("0.4223", "percentage")}},
	"formula":                {{MakeCell("A1+B1", "formula")}},
	"currency default (eur)": {{MakeCell("2.22", "currency")}},
	"currency eur negative":  {{MakeCell("-2.22", "currency-eur")}},
	"currency usd":           {{MakeCell("2.22", "currency-usd")}},
	"currency usd negative":  {{MakeCell("-2.22", "currency-usd")}},
	"currency gbp":           {{MakeCell("2.22", "currency-gbp")}},
	"currency gbp negative":  {{MakeCell("-2.22", "currency-gbp")}},
	"named range":            {{MakeRangeCell("42", "float", "answer")}},
	"styled cell": {{
		MakeStyledCell("Navy", "string", CellStyle{BackgroundColor: "#001f3f"}),
		MakeStyledCell("42.33", "float", CellStyle{Bold: true, Italic: true, FontColor: "#ffffff", Border: "0.5pt solid #000000"}),
	}},
	"all types combined": {
		{
			MakeCell("ABBA", "string"),
			MakeCell("42.3324", "float"),
			MakeCell("-42.3324", "float"),
			MakeCell("2022-02-02", "date"),
			MakeCell("19:03:00", "time"),
			MakeCell("2.22", "currency"),
			MakeCell("2.22", "currency-usd"),
			MakeCell("2.22", "currency-gbp"),
			MakeCell("0.4223", "percentage"),
			MakeCell("A1+B1", "formula"),
			MakeRangeCell("42", "float", "answer"),
		},
	},
}

func TestOdsPartsMatchOdfSchema(t *testing.T) {
	for name, cells := range schemaTestCases {
		t.Run(name, func(t *testing.T) {
			spreadsheet, err := MakeSpreadsheet(cells)
			if err != nil {
				t.Fatalf("MakeSpreadsheet: %v", err)
			}

			buff, err := MakeOds(spreadsheet)
			if err != nil {
				t.Fatalf("MakeOds: %v", err)
			}

			reader, err := zip.NewReader(bytes.NewReader(buff.Bytes()), int64(buff.Len()))
			if err != nil {
				t.Fatal(err)
			}

			parts := map[string]string{}
			for _, f := range reader.File {
				rc, err := f.Open()
				if err != nil {
					t.Fatal(err)
				}
				content, err := io.ReadAll(rc)
				if err != nil {
					t.Fatal(err)
				}
				if err := rc.Close(); err != nil {
					t.Fatal(err)
				}
				parts[f.Name] = string(content)
			}

			for _, partName := range []string{"content.xml", "styles.xml", "meta.xml"} {
				t.Run(partName, func(t *testing.T) {
					content, ok := parts[partName]
					if !ok {
						t.Fatalf("package does not contain %s", partName)
					}
					validateAgainstSchema(t, partName, content)
				})
			}
		})
	}
}

func TestAutoFilterMatchesOdfSchema(t *testing.T) {
	spreadsheet, err := MakeSpreadsheet([][]Cell{
		{MakeCell("Name", "string"), MakeCell("Age", "string")},
		{MakeCell("Alice", "string"), MakeCell("30", "float")},
	})
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}
	spreadsheet = EnableAutoFilter(spreadsheet)

	flatOds, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}
	validateAgainstSchema(t, "flat.fods", flatOds)
}

func TestTableMatchesOdfSchema(t *testing.T) {
	spreadsheet, err := MakeTable([][]Cell{
		{MakeCell("Product", "string"), MakeCell("Price", "string")},
		{MakeCell("Pen", "string"), MakeCell("1.49", "float")},
		{MakeCell("Desk", "string"), MakeCell("189.00", "float")},
	}, TableOptions{
		Name:       "Products",
		Header:     true,
		AutoFilter: true,
		BandedRows: true,
		Totals:     []Total{{TotalNone}, {TotalSum}},
	})
	if err != nil {
		t.Fatalf("MakeTable: %v", err)
	}

	flatOds, err := MakeFlatOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeFlatOds: %v", err)
	}
	validateAgainstSchema(t, "flat.fods", flatOds)
}

func TestFlatOdsMatchesOdfSchema(t *testing.T) {
	for name, cells := range schemaTestCases {
		t.Run(name, func(t *testing.T) {
			spreadsheet, err := MakeSpreadsheet(cells)
			if err != nil {
				t.Fatalf("MakeSpreadsheet: %v", err)
			}

			flatOds, err := MakeFlatOds(spreadsheet)
			if err != nil {
				t.Fatalf("MakeFlatOds: %v", err)
			}

			validateAgainstSchema(t, "flat.fods", flatOds)
		})
	}
}
