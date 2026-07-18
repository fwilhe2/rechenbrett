<!--
SPDX-FileCopyrightText: 2025 Florian Wilhelm

SPDX-License-Identifier: MIT
-->

# rechenbrett

A go library for building Open Document spreadsheet files.

It can create 'normal' ods (Open Document Spreadsheet) files (`*.ods`) and 'flat' Open Document Spreadsheet files (`*.fods`).

`ods` files are zipped and can be opened with various commercial and open source spreadsheet applications.
This is the default file format used by LibreOffice Calc when saving data.

`fods` files are plain xml files without compression.
They contain the same information than their zipped counterparts, but wrap everything in one large xml document.
Due to their plain text nature, they work well with version control systems such as git.
For example, if you want to keep track of your bank account statements, which you might get in some sort of complex xml or json structure, you could use rechenbrett to convert them into a clean flat ods structure which can be version controlled and produce meaningful diffs.

Sadly, if you save `fods` files using LibreOffice Calc, it changes the file in many places which makes it harder to diff two versions of the same file in a meaningful way.
Post processing the file using [`flat-odf-cleanup.py`](https://github.com/fwilhe2/odf-utils/blob/main/flat-odf-cleanup.py) can mitigate the issue, but does not fully resolve it.

## Example usage

```go
package main

import (
	"os"

	rb "github.com/fwilhe2/rechenbrett"
)

func main() {
	inputCells := [][]rb.Cell{
		{
			rb.MakeCell("ABBA", "string"),
			rb.MakeCell("42.3324", "float"),
			rb.MakeCell("2022-02-02", "date"),
			rb.MakeCell("2.2.2022", "date"),
			rb.MakeCell("19:03:00", "time"),
			rb.MakeCell("2.22", "currency"),
			rb.MakeCell("-2.22", "currency"),
			rb.MakeCell("0.4223", "percentage"),
		},
	}

	spreadsheet, err := rb.MakeSpreadsheet(inputCells)
	if err != nil {
		panic(err)
	}

	// create ods file
	buff, err := rb.MakeOds(spreadsheet)
	if err != nil {
		panic(err)
	}
	if err := os.WriteFile("myfile.ods", buff.Bytes(), 0o644); err != nil {
		panic(err)
	}

	// create fods file
	flatOdsString, err := rb.MakeFlatOds(spreadsheet)
	if err != nil {
		panic(err)
	}
	if err := os.WriteFile("myfile.fods", []byte(flatOdsString), 0o644); err != nil {
		panic(err)
	}
}
```

## API surface

Cells are created with `MakeCell` or `MakeRangeCell`, arranged into rows, combined into a `Spreadsheet` with `MakeSpreadsheet`, and serialized with `MakeOds`, `WriteOds`, or `MakeFlatOds`.

- `MakeCell(value, valueType string) Cell` — creates a cell holding `value` interpreted as `valueType`. Supported value types:
  - `"string"`
  - `"float"`
  - `"date"` (ISO `YYYY-MM-DD`, German `DD.MM.YYYY`, or US `MM/DD/YYYY`)
  - `"time"` (`HH:MM` or `HH:MM:SS`)
  - `"percentage"` (a fraction, e.g. `"0.42"` for 42 %)
  - `"formula"`
  - `"currency"` (defaults to EUR), `"currency-eur"`, `"currency-usd"`, `"currency-gbp"`

  Invalid values or value types are not reported here; they surface as an error from `MakeSpreadsheet`.

- `MakeRangeCell(value, valueType, rangeName string) Cell` — like `MakeCell`, and additionally names the cell's position as `rangeName` so formulas in other cells can refer to it by name. Each range name may be used for only one cell.

- `MakeSpreadsheet(cells [][]Cell) (Spreadsheet, error)` — arranges the given rows of cells into a spreadsheet with a single sheet named `Sheet1`. Reports all invalid cells (bad value types, unparseable dates/times/numbers) and duplicate range names together as a single joined error.

- `MakeSpreadsheetWithName(name string, cells [][]Cell) (Spreadsheet, error)` — like `MakeSpreadsheet`, with a custom sheet name.

- `MakeOds(spreadsheet Spreadsheet) (*bytes.Buffer, error)` — serializes the spreadsheet as a zipped OpenDocument package (`.ods`).

- `WriteOds(w io.Writer, spreadsheet Spreadsheet) error` — like `MakeOds`, but streams the zipped package directly to `w`.

- `MakeFlatOds(spreadsheet Spreadsheet) (string, error)` — serializes the spreadsheet as a flat OpenDocument XML document (`.fods`).

`Cell` and `Spreadsheet` are the only exported types beyond the functions above; their fields are exported solely for XML marshaling and aren't meant to be constructed or read directly — build values through the functions instead.

## Related

[json-to-ods](https://github.com/fwilhe2/mkods) is a simple go wrapper for rechenbrett to make it usable as a cli tool

[mkods-demo](https://github.com/fwilhe2/mkods-demo) shows how *mkods* can be used in combination with node.js to transform complex json structures into clean spreadsheets

[csv-to-ods](https://github.com/fwilhe2/csv-to-ods) converts csv files to ods with optional type hints

[kalkulationsbogen](https://github.com/fwilhe2/kalkulationsbogen) is similar to rechenbrett, but written in TypeScript for node.js

## License

This software is written by Florian Wilhelm and available under the MIT license (see `LICENSE` for details)
