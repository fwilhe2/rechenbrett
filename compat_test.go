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
	styleDefinition   = regexp.MustCompile(`style:name="([^"]+)"`)
	parentStyleUse    = regexp.MustCompile(`style:parent-style-name="([^"]+)"`)
	dataStyleUse      = regexp.MustCompile(`style:data-style-name="([^"]+)"`)
	cellStyleUse      = regexp.MustCompile(`table:style-name="([^"]+)"`)
	masterPageUse     = regexp.MustCompile(`style:master-page-name="([^"]+)"`)
	pageLayoutUse     = regexp.MustCompile(`style:page-layout-name="([^"]+)"`)
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

// TestCompatStyleReferencesResolve checks that every style a document refers
// to is defined somewhere in it. A dangling reference costs the formatting
// attached to it in consumers that, unlike LibreOffice, do not invent the
// missing style.
func TestCompatStyleReferencesResolve(t *testing.T) {
	parts := readOdsParts(t, compatSpreadsheet(t))
	both := parts["content.xml"] + parts["styles.xml"]

	defined := map[string]bool{}
	for _, name := range matches(styleDefinition, both) {
		defined[name] = true
	}

	for _, pattern := range []*regexp.Regexp{
		parentStyleUse, dataStyleUse, cellStyleUse, masterPageUse, pageLayoutUse,
	} {
		for _, used := range matches(pattern, both) {
			if !defined[used] {
				t.Errorf("style %q is referenced but never defined", used)
			}
		}
	}
}

// TestCompatStylesXmlHasPageSetup checks that styles.xml carries the common
// styles and the master page. Excel expects a page setup for every sheet.
func TestCompatStylesXmlHasPageSetup(t *testing.T) {
	parts := readOdsParts(t, compatSpreadsheet(t))
	styles := parts["styles.xml"]

	for _, expected := range []string{
		"<office:styles>",
		"<office:master-styles>",
		"<style:master-page",
		"<style:page-layout",
	} {
		if !strings.Contains(styles, expected) {
			t.Errorf("styles.xml does not contain %s", expected)
		}
	}

	if !strings.Contains(parts["content.xml"], `table:style-name="`+tableStyleName+`"`) {
		t.Errorf("the sheet does not refer to the table style binding it to the master page")
	}
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

// TestCompatCurrencyStyleIsNotVolatile checks that the currency styles cells
// refer to are not marked volatile. A volatile style is one a consumer may
// discard, which would take the currency formatting with it. Only the styles
// reached through style:map are volatile.
func TestCompatCurrencyStyleIsNotVolatile(t *testing.T) {
	content := readOdsParts(t, compatSpreadsheet(t))["content.xml"]

	for _, code := range []string{"EUR", "USD", "GBP"} {
		definition := regexp.MustCompile(`<number:currency-style style:name="` + code + `_DATA_STYLE"([^>]*)>`)
		match := definition.FindStringSubmatch(content)
		if match == nil {
			t.Fatalf("%s_DATA_STYLE is not defined", code)
		}
		if strings.Contains(match[1], "style:volatile") {
			t.Errorf("%s_DATA_STYLE is marked volatile, but cells refer to it", code)
		}
	}
}

// TestCompatZipTimestampsAreValid checks that no entry of the package
// predates 1980. Earlier dates cannot be expressed in the MS-DOS date fields
// of a zip archive, and strict readers reject an archive carrying them.
func TestCompatZipTimestampsAreValid(t *testing.T) {
	buff, err := MakeOds(compatSpreadsheet(t))
	if err != nil {
		t.Fatalf("MakeOds: %v", err)
	}
	reader, err := zip.NewReader(bytes.NewReader(buff.Bytes()), int64(buff.Len()))
	if err != nil {
		t.Fatal(err)
	}

	for _, f := range reader.File {
		if f.Modified.Year() < 1980 {
			t.Errorf("zip entry %s is dated %s, which a zip archive cannot express", f.Name, f.Modified)
		}
	}
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
		t.Skip("ssconvert (gnumeric) is not installed, skipping cross-reader test")
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
