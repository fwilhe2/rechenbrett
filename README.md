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

	spreadsheet := rb.MakeSpreadsheet(inputCells)

	// create ods file
	buff := rb.MakeOds(spreadsheet)
	archive, err := os.Create("myfile.ods")
	if err != nil {
		panic(err)
	}
	archive.Write(buff.Bytes())
	archive.Close()

	// create fods file
	flatOdsString := rb.MakeFlatOds(spreadsheet)
	os.WriteFile("myfile.fods", []byte(flatOdsString), 0o644)
}
```

## Related

[json-to-ods](https://github.com/fwilhe2/mkods) is a simple go wrapper for rechenbrett to make it usable as a cli tool

[mkods-demo](https://github.com/fwilhe2/mkods-demo) shows how *mkods* can be used in combination with node.js to transform complex json structures into clean spreadsheets

[csv-to-ods](https://github.com/fwilhe2/csv-to-ods) converts csv files to ods with optional type hints

[kalkulationsbogen](https://github.com/fwilhe2/kalkulationsbogen) is similar to rechenbrett, but written in TypeScript for node.js

## License

This software is written by Florian Wilhelm and available under the MIT license (see `LICENSE` for details)
