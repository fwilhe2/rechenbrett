// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

// Command showcase generates example .ods and .fods documents that exercise
// rechenbrett's features, for eyeballing in a spreadsheet application or
// spot-checking output without running the full (LibreOffice-backed) test
// suite.
//
//	go run ./cmd/showcase [-out output]
package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	rb "github.com/fwilhe2/rechenbrett"
)

func main() {
	out := flag.String("out", "output", "directory to write generated documents to")
	flag.Parse()

	if err := os.MkdirAll(*out, 0o755); err != nil {
		log.Fatal(err)
	}

	documents := map[string]rb.Spreadsheet{
		"data-types":  mustSpreadsheet("data-types", dataTypesDocument()),
		"styles":      mustSpreadsheet("styles", stylesDocument()),
		"auto-filter": autoFilterDocument(),
	}

	for name, spreadsheet := range documents {
		odsPath := filepath.Join(*out, name+".ods")
		odsFile, err := os.Create(odsPath)
		if err != nil {
			log.Fatal(err)
		}
		if err := rb.WriteOds(odsFile, spreadsheet); err != nil {
			log.Fatalf("%s: %v", name, err)
		}
		if err := odsFile.Close(); err != nil {
			log.Fatal(err)
		}
		fmt.Println("wrote", odsPath)

		fods, err := rb.MakeFlatOds(spreadsheet)
		if err != nil {
			log.Fatalf("%s: %v", name, err)
		}
		fodsPath := filepath.Join(*out, name+".fods")
		if err := os.WriteFile(fodsPath, []byte(fods), 0o644); err != nil {
			log.Fatal(err)
		}
		fmt.Println("wrote", fodsPath)
	}
}

// mustSpreadsheet builds a spreadsheet from cells, exiting on error. The
// showcase inputs are all valid, so an error here is a programming mistake.
func mustSpreadsheet(name string, cells [][]rb.Cell) rb.Spreadsheet {
	spreadsheet, err := rb.MakeSpreadsheet(cells)
	if err != nil {
		log.Fatalf("%s: %v", name, err)
	}
	return spreadsheet
}

// dataTypesDocument exercises every value type MakeCell supports, plus
// formulas referring to a named range.
func dataTypesDocument() [][]rb.Cell {
	return [][]rb.Cell{
		{
			rb.MakeCell("String", "string"),
			rb.MakeCell("Float", "string"),
			rb.MakeCell("Date", "string"),
			rb.MakeCell("Time", "string"),
			rb.MakeCell("Percentage", "string"),
			rb.MakeCell("EUR", "string"),
			rb.MakeCell("USD", "string"),
			rb.MakeCell("GBP", "string"),
			rb.MakeCell("Formula", "string"),
		},
		{
			rb.MakeCell("ABBA", "string"),
			rb.MakeRangeCell("42.3324", "float", "ShowcaseFloat"),
			rb.MakeCell("2022-02-02", "date"),
			rb.MakeCell("19:03:00", "time"),
			rb.MakeCell("0.4223", "percentage"),
			rb.MakeCell("2.22", "currency-eur"),
			rb.MakeCell("2.22", "currency-usd"),
			rb.MakeCell("2.22", "currency-gbp"),
			rb.MakeCell("ShowcaseFloat*2", "formula"),
		},
	}
}

// stylesDocument exercises MakeStyledCell and the built-in Color palette
// with a small header-row-style table.
func stylesDocument() [][]rb.Cell {
	header := rb.CellStyle{
		BackgroundColor: rb.ColorNavy,
		FontColor:       rb.ColorWhite,
		Bold:            true,
		Border:          "0.5pt solid #000000",
	}

	palette := []struct {
		name  string
		color string
	}{
		{"Navy", rb.ColorNavy},
		{"Blue", rb.ColorBlue},
		{"Aqua", rb.ColorAqua},
		{"Teal", rb.ColorTeal},
		{"Purple", rb.ColorPurple},
		{"Fuchsia", rb.ColorFuchsia},
		{"Maroon", rb.ColorMaroon},
		{"Red", rb.ColorRed},
		{"Orange", rb.ColorOrange},
		{"Yellow", rb.ColorYellow},
		{"Olive", rb.ColorOlive},
		{"Green", rb.ColorGreen},
		{"Lime", rb.ColorLime},
		{"Black", rb.ColorBlack},
		{"Gray", rb.ColorGray},
		{"Silver", rb.ColorSilver},
		{"White", rb.ColorWhite},
	}

	rows := [][]rb.Cell{
		{rb.MakeStyledCell("Color palette (clrs.cc)", "string", header)},
	}
	for _, c := range palette {
		rows = append(rows, []rb.Cell{
			rb.MakeStyledCell(c.name, "string", rb.CellStyle{BackgroundColor: c.color}),
		})
	}
	return rows
}

// autoFilterDocument shows EnableAutoFilter: a small table that opens with
// AutoFilter dropdown buttons on its header row.
func autoFilterDocument() rb.Spreadsheet {
	header := rb.CellStyle{
		BackgroundColor: rb.ColorNavy,
		FontColor:       rb.ColorWhite,
		Bold:            true,
	}
	cells := [][]rb.Cell{
		{
			rb.MakeStyledCell("Product", "string", header),
			rb.MakeStyledCell("Category", "string", header),
			rb.MakeStyledCell("Price", "string", header),
		},
		{rb.MakeCell("Laptop", "string"), rb.MakeCell("Electronics", "string"), rb.MakeCell("999.00", "currency")},
		{rb.MakeCell("Mouse", "string"), rb.MakeCell("Electronics", "string"), rb.MakeCell("19.99", "currency")},
		{rb.MakeCell("Desk", "string"), rb.MakeCell("Furniture", "string"), rb.MakeCell("189.00", "currency")},
		{rb.MakeCell("Pen", "string"), rb.MakeCell("Stationery", "string"), rb.MakeCell("1.49", "currency")},
	}

	return rb.EnableAutoFilter(mustSpreadsheet("auto-filter", cells))
}
