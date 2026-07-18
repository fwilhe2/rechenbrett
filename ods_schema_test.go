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

// schemaTestSpreadsheet covers every cell type the library can produce.
func schemaTestSpreadsheet() Spreadsheet {
	return MakeSpreadsheet([][]Cell{
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
	})
}

func TestOdsPartsMatchOdfSchema(t *testing.T) {
	buff := MakeOds(schemaTestSpreadsheet())

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

	for _, name := range []string{"content.xml", "styles.xml", "meta.xml"} {
		t.Run(name, func(t *testing.T) {
			content, ok := parts[name]
			if !ok {
				t.Fatalf("package does not contain %s", name)
			}
			validateAgainstSchema(t, name, content)
		})
	}
}

func TestFlatOdsMatchesOdfSchema(t *testing.T) {
	validateAgainstSchema(t, "flat.fods", MakeFlatOds(schemaTestSpreadsheet()))
}
