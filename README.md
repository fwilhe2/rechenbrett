# rechenbrett

A go library for building Open Document spreadsheet files.

It can create 'normal' ods (Open Document Spreadsheet) files (`*.ods`) and 'flat' ods files (`*.fods`).

`ods` files are zipped and can be opened with various commercial and open source spreadsheet applications.

`fods` files are plain xml files without compression.
Due to their plain text nature, they work well with version control systems such as git.
For example, if you want to keep track of your bank account statements, which you might get in some sort of complex xml or json structure, you could use rechenbrett to convert them into a clean flat ods structure which can be version controlled and produce meaningful diffs.

## Related

[mkods](https://github.com/fwilhe2/mkods) is a simple go wrapper for rechenbrett to make it usable as a cli tool

[mkods-demo](https://github.com/fwilhe2/mkods-demo) shows how *mkods* can be used in combination with node.js to transform complex json structures into clean spreadsheets

[kalkulationsbogen](https://github.com/fwilhe2/kalkulationsbogen) is similar to rechenbrett, but written in TypeScript for node.js

## License

This software is written by Florian Wilhelm and available under the MIT license (see `LICENSE` for details)
