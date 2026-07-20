// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// The tests in this file guard against output that LibreOffice reads but
// stricter consumers, Excel above all, do not. LibreOffice repairs a great
// deal silently, so the integration tests alone do not notice when the
// documents stop conforming.

// readOdsParts unzips a generated package into a map of part name to content.
func readOdsParts(t *testing.T, spreadsheet Spreadsheet) map[string]string {
	t.Helper()

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
	return parts
}

// compatSpreadsheet exercises the features whose serialization the tests
// below inspect.
func compatSpreadsheet(t *testing.T) Spreadsheet {
	t.Helper()

	spreadsheet, err := MakeSpreadsheet([][]Cell{
		{
			MakeRangeCell("42.3324", "float", "InputA"),
			MakeCell("23", "float"),
			MakeCell("2022-02-02", "date"),
			MakeCell("2.22", "currency"),
			MakeCell("SUM(A1:B1)", "formula"),
			MakeStyledCell("Navy", "string", CellStyle{BackgroundColor: ColorNavy}),
		},
	})
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}
	return spreadsheet
}

var (
	formulaAttribute  = regexp.MustCompile(`table:formula="([^"]+)"`)
	namespaceQualifie = regexp.MustCompile(`^[a-z]+:=`)
)

func matches(pattern *regexp.Regexp, content string) []string {
	var found []string
	for _, m := range pattern.FindAllStringSubmatch(content, -1) {
		found = append(found, m[1])
	}
	return found
}

// TestCompatFormulasAreNamespaced checks that stored formulas carry a
// namespace prefix and bracket their cell references. Without either, a
// conforming consumer drops the formula or misreads its references.
func TestCompatFormulasAreNamespaced(t *testing.T) {
	parts := readOdsParts(t, compatSpreadsheet(t))

	formulas := matches(formulaAttribute, parts["content.xml"])
	if len(formulas) == 0 {
		t.Fatal("expected the document to contain a formula")
	}
	for _, formula := range formulas {
		if !namespaceQualifie.MatchString(formula) {
			t.Errorf("formula %q has no namespace prefix", formula)
		}
		if strings.Contains(formula, "A1") && !strings.Contains(formula, "[.A1") {
			t.Errorf("formula %q does not bracket its cell references", formula)
		}
	}

	if !strings.Contains(parts["content.xml"], `xmlns:of="`) {
		t.Error("content.xml does not declare the OpenFormula namespace")
	}
}

// crossReadersRequired reports whether a missing cross-reader is a failure
// rather than a reason to skip. The dedicated CI jobs set it, so that a
// cross-reader that is not installed, or an image that cannot be pulled,
// fails visibly instead of quietly testing nothing.
func crossReadersRequired() bool {
	return os.Getenv("RECHENBRETT_REQUIRE_CROSS_READERS") != ""
}

// missingCrossReader skips the calling test, or fails it when the
// cross-readers are required.
func missingCrossReader(t *testing.T, format string, args ...any) {
	t.Helper()
	if crossReadersRequired() {
		t.Fatalf(format, args...)
	}
	t.Skipf(format, args...)
}

// TestCompatGnumericReadsDocument has Gnumeric convert a generated document
// to CSV and compares the result with what the document should contain.
//
// Gnumeric is a second implementation of the format, independent of
// LibreOffice and considerably stricter about it. It stands in for the
// consumers this package cannot be tested against directly: what Gnumeric
// silently drops or misreads, Excel tends to drop or misread too.
func TestCompatGnumericReadsDocument(t *testing.T) {
	ssconvert, err := exec.LookPath("ssconvert")
	if err != nil {
		missingCrossReader(t, "ssconvert (gnumeric) is not installed")
	}

	cells := [][]Cell{
		{
			MakeRangeCell("42.3324", "float", "InputA"),
			MakeCell("23", "float"),
			MakeCell("A1+B1", "formula"),
			MakeCell("SUM(A1:B1)", "formula"),
			MakeCell("InputA*2", "formula"),
			MakeCell("ABBA", "string"),
		},
	}
	// Gnumeric writes the unformatted values, so the expectation is the
	// values themselves rather than the locale-specific rendering the
	// LibreOffice integration tests compare against.
	expected := []string{"42.3324", "23", "65.3324", "65.3324", "84.6648", "ABBA"}

	spreadsheet, err := MakeSpreadsheet(cells)
	if err != nil {
		t.Fatalf("MakeSpreadsheet: %v", err)
	}
	buff, err := MakeOds(spreadsheet)
	if err != nil {
		t.Fatalf("MakeOds: %v", err)
	}

	dir := t.TempDir()
	odsPath := filepath.Join(dir, "gnumeric.ods")
	csvPath := filepath.Join(dir, "gnumeric.csv")
	if err := os.WriteFile(odsPath, buff.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}

	if output, err := exec.Command(ssconvert, odsPath, csvPath).CombinedOutput(); err != nil {
		t.Fatalf("gnumeric conversion failed: %v\n%s", err, output)
	}

	converted, err := os.ReadFile(csvPath)
	if err != nil {
		t.Fatal(err)
	}
	record, err := csv.NewReader(bytes.NewReader(converted)).Read()
	if err != nil {
		t.Fatalf("reading the converted csv: %v", err)
	}

	if len(record) != len(expected) {
		t.Fatalf("gnumeric read %d cells, expected %d: %q", len(record), len(expected), record)
	}
	for i, want := range expected {
		if record[i] != want {
			t.Errorf("cell %d: gnumeric read %q, expected %q", i+1, record[i], want)
		}
	}
}
