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

	// create ods file: prefer WriteOds over MakeOds when writing to a file or
	// other io.Writer, since it streams the zip archive directly instead of
	// building it in a buffer first
	archive, err := os.Create("myfile.ods")
	if err != nil {
		panic(err)
	}
	defer archive.Close()
	if err := rb.WriteOds(archive, spreadsheet); err != nil {
		panic(err)
	}

	// create fods file: MakeFlatOds always builds the whole document in
	// memory (xml.MarshalIndent has no streaming variant to wrap), so there
	// is no WriteFods equivalent to WriteOds — just write out the string
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

Cells are created with `MakeCell`, `MakeRangeCell`, or `MakeStyledCell`, arranged into rows, combined into a `Spreadsheet` with `MakeSpreadsheet`, and serialized with `MakeOds`, `WriteOds`, or `MakeFlatOds`.

- `MakeCell(value, valueType string) Cell` — creates a cell holding `value` interpreted as `valueType`. Supported value types:
  - `"string"`
  - `"float"`
  - `"date"` (ISO `YYYY-MM-DD`, German `DD.MM.YYYY`, or US `MM/DD/YYYY`)
  - `"time"` (`HH:MM` or `HH:MM:SS`)
  - `"percentage"` (a fraction, e.g. `"0.42"` for 42 %)
  - `"formula"` (in the familiar A1 notation, e.g. `"SUM(A1:B1)"`, `"InputA*2"`, without a leading `=`; it is translated to the OpenFormula notation the format stores, e.g. `of:=SUM([.A1:.B1])`)
  - `"currency"` (defaults to EUR), `"currency-eur"`, `"currency-usd"`, `"currency-gbp"`

  Invalid values or value types are not reported here; they surface as an error from `MakeSpreadsheet`.

- `MakeRangeCell(value, valueType, rangeName string) Cell` — like `MakeCell`, and additionally names the cell's position as `rangeName` so formulas in other cells can refer to it by name. Each range name may be used for only one cell.

- `MakeStyledCell(value, valueType string, style CellStyle) Cell` — like `MakeCell`, and additionally applies `style` to the cell's appearance. `CellStyle` sets `BackgroundColor` and `FontColor` (hex strings, e.g. `"#ff0000"`), `Bold`/`Italic`, and `Border` (an ODF `fo:border` shorthand value, e.g. `"0.5pt solid #000000"`, applied to all four sides). Cells created with an identical style share a single generated style definition.

  `Color*` constants (`ColorNavy`, `ColorBlue`, `ColorAqua`, `ColorTeal`, `ColorPurple`, `ColorFuchsia`, `ColorMaroon`, `ColorRed`, `ColorOrange`, `ColorYellow`, `ColorOlive`, `ColorGreen`, `ColorLime`, `ColorBlack`, `ColorGray`, `ColorSilver`, `ColorWhite`), taken from the palette at [clrs.cc](https://clrs.cc/), are available for use as `BackgroundColor`/`FontColor` values.

- `MakeSpreadsheet(cells [][]Cell) (Spreadsheet, error)` — arranges the given rows of cells into a spreadsheet with a single sheet named `Sheet1`. Reports all invalid cells (bad value types, unparseable dates/times/numbers) and duplicate range names together as a single joined error.

- `MakeSpreadsheetWithName(name string, cells [][]Cell) (Spreadsheet, error)` — like `MakeSpreadsheet`, with a custom sheet name.

- `EnableAutoFilter(spreadsheet Spreadsheet) Spreadsheet` — returns the spreadsheet with AutoFilter dropdown buttons enabled over the used cell range of every non-empty sheet, so the generated document opens with filter dropdowns on the header row. It sets the buttons only (no saved filter conditions, so all rows stay visible); calling it again replaces any previously enabled AutoFilter. Compose it with the `MakeSpreadsheet` result before serializing:

  ```go
  spreadsheet, err := rb.MakeSpreadsheet(cells)
  // ...
  spreadsheet = rb.EnableAutoFilter(spreadsheet)
  ```

- `MakeOds(spreadsheet Spreadsheet) (*bytes.Buffer, error)` — serializes the spreadsheet as a zipped OpenDocument package (`.ods`). Implemented as `WriteOds` into a `bytes.Buffer`; prefer calling `WriteOds` directly when you already have an `io.Writer` (a file, an HTTP response, ...) to avoid the extra buffer copy.

- `WriteOds(w io.Writer, spreadsheet Spreadsheet) error` — writes the zipped OpenDocument package (`.ods`) directly to `w`. This is the recommended entry point for producing `.ods` output: it streams archive entries straight to `w` via `archive/zip`, rather than materializing the whole archive in memory first.

- `MakeFlatOds(spreadsheet Spreadsheet) (string, error)` — serializes the spreadsheet as a flat OpenDocument XML document (`.fods`). There is no `WriteFods` counterpart to `WriteOds`: the flat document is built with `xml.MarshalIndent`, which has no streaming variant, so the full document is always materialized in memory before `MakeFlatOds` returns it as a string — a `Write` variant would offer no benefit over calling `MakeFlatOds` and writing the result yourself.

`Cell`, `Spreadsheet`, and `CellStyle` are the only exported types beyond the functions above. `Cell` and `Spreadsheet` fields are exported solely for XML marshaling and aren't meant to be constructed or read directly — build values through the functions instead.

## Showcase

`make showcase` (or `go run ./cmd/showcase`) generates example `.ods` and `.fods` documents into `output/` (gitignored) that exercise rechenbrett's features — every value type, formulas and named ranges, custom cell styles with the `Color*` palette, and an AutoFilter table — for opening in a spreadsheet application or spot-checking output. It runs in well under a second and needs no LibreOffice install, unlike the test suite (`make test`), which drives LibreOffice to verify rendered values.

## Compatibility with other spreadsheet applications

The generated documents are validated against the OpenDocument schema and are meant to open in any application that reads the format, not just LibreOffice.

That is not something the LibreOffice-backed tests can establish on their own: LibreOffice accepts a good deal that the format does not actually allow, and repairs the rest silently, so documents that it renders correctly can still be misread elsewhere. Two things guard against this:

- `compat_test.go` inspects the generated XML directly for the mistakes stricter consumers punish — dangling style references, formulas without a namespace prefix, a missing page setup, styles marked as discardable, zip timestamps a zip archive cannot express.
- The same file has [Gnumeric](http://www.gnumeric.org/) read a generated document back. Gnumeric is an implementation of the format independent of LibreOffice and considerably stricter about it, which makes it a usable stand-in for the applications this package cannot be tested against directly. The test skips itself when `ssconvert` is not installed.

Testing against Microsoft Excel itself is not automated: it requires either a Windows machine with a licensed Office driving Excel through COM, or a Microsoft 365 account and the Graph API to have Excel's own import engine open the file server-side.

## Related

[json-to-ods](https://github.com/fwilhe2/mkods) is a simple go wrapper for rechenbrett to make it usable as a cli tool

[mkods-demo](https://github.com/fwilhe2/mkods-demo) shows how *mkods* can be used in combination with node.js to transform complex json structures into clean spreadsheets

[csv-to-ods](https://github.com/fwilhe2/csv-to-ods) converts csv files to ods with optional type hints

[kalkulationsbogen](https://github.com/fwilhe2/kalkulationsbogen) is similar to rechenbrett, but written in TypeScript for node.js

## License

This software is written by Florian Wilhelm and available under the MIT license (see `LICENSE` for details)
