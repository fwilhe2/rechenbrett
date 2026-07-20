// SPDX-FileCopyrightText: 2025 Florian Wilhelm
//
// SPDX-License-Identifier: MIT

package ods

import (
	"regexp"
	"strings"
)

// formulaNamespace is the OpenFormula namespace prefix that table:formula
// values are expected to carry.
const formulaNamespace = "of:"

// cellReference matches a single A1-style cell address, with optional
// absolute markers, as written by callers of MakeCell.
var cellReference = regexp.MustCompile(`^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$`)

// toOpenFormula converts the A1-style formula callers write ("SUM(A1:B1)")
// into the OpenFormula expression stored in table:formula
// ("of:=SUM([.A1:.B1])").
//
// Three things differ between the two notations: cell references are enclosed
// in square brackets and prefixed with a sheet separator, function arguments
// are separated by semicolons rather than commas, and the whole expression
// carries a namespace prefix. LibreOffice accepts the plain notation, but
// conforming consumers (Excel, Gnumeric) either drop such formulas or, worse,
// misread the references and compute wrong results.
//
// Identifiers that are not cell addresses are left alone: function names are
// recognized by the following parenthesis, everything else is assumed to be a
// named range and referenced by name, which is how OpenFormula spells it too.
//
// A formula that already carries a namespace prefix is passed through
// unchanged, as an escape hatch for expressions this translation cannot
// express.
func toOpenFormula(formula string) string {
	if strings.Contains(formula, ":=") {
		return formula
	}
	formula = strings.TrimPrefix(formula, "=")

	tokens := tokenizeFormula(formula)
	var b strings.Builder
	b.WriteString(formulaNamespace + "=")
	for i := 0; i < len(tokens); i++ {
		t := tokens[i]
		switch {
		case t.kind == tokenIdent && isFunctionCall(tokens, i):
			b.WriteString(t.text)
		case t.kind == tokenIdent && isReference(t.text):
			// A reference followed by ":" and a second reference is a range,
			// which OpenFormula writes as a single bracketed expression.
			if end, ok := rangeEnd(tokens, i); ok {
				b.WriteString("[" + bracketBody(t.text) + ":" + bracketBody(tokens[end].text) + "]")
				i = end
				continue
			}
			b.WriteString("[" + bracketBody(t.text) + "]")
		case t.kind == tokenOther && t.text == ",":
			b.WriteString(";")
		default:
			b.WriteString(t.text)
		}
	}
	return b.String()
}

// bracketBody renders a reference for use inside square brackets: an
// unqualified address is prefixed with "." to denote the current sheet,
// a sheet-qualified one ("Sheet1.A1") is already in the expected form.
func bracketBody(reference string) string {
	if strings.Contains(reference, ".") {
		return reference
	}
	return "." + reference
}

// isReference reports whether an identifier is a cell address, either
// unqualified ("A1") or sheet-qualified ("Sheet1.A1"). Anything else is a
// named range.
func isReference(identifier string) bool {
	if dot := strings.LastIndex(identifier, "."); dot >= 0 {
		return dot > 0 && cellReference.MatchString(identifier[dot+1:])
	}
	return cellReference.MatchString(identifier)
}

// isFunctionCall reports whether the identifier at index i is followed by an
// opening parenthesis, which makes it a function name rather than a
// reference.
func isFunctionCall(tokens []token, i int) bool {
	next := nextSignificant(tokens, i+1)
	return next >= 0 && tokens[next].text == "("
}

// rangeEnd reports the index of the closing reference of a range starting at
// index i, if the reference at i is followed by ":" and a second reference.
func rangeEnd(tokens []token, i int) (int, bool) {
	colon := nextSignificant(tokens, i+1)
	if colon < 0 || tokens[colon].text != ":" {
		return 0, false
	}
	end := nextSignificant(tokens, colon+1)
	if end < 0 || tokens[end].kind != tokenIdent || !isReference(tokens[end].text) {
		return 0, false
	}
	return end, true
}

// nextSignificant returns the index of the first token at or after i that is
// not whitespace, or -1 if there is none.
func nextSignificant(tokens []token, i int) int {
	for ; i < len(tokens); i++ {
		if strings.TrimSpace(tokens[i].text) != "" {
			return i
		}
	}
	return -1
}

type tokenKind int

const (
	// tokenOther covers operators, parentheses, numbers, and whitespace:
	// everything that is copied through unchanged except for the comma.
	tokenOther tokenKind = iota
	tokenIdent
	tokenString
)

type token struct {
	kind tokenKind
	text string
}

// tokenizeFormula splits a formula into identifiers, string literals, and
// single characters. It is deliberately loose: anything it does not
// recognize ends up as a tokenOther and is copied through verbatim.
func tokenizeFormula(formula string) []token {
	var tokens []token
	runes := []rune(formula)
	for i := 0; i < len(runes); {
		c := runes[i]
		switch {
		case c == '"':
			// String literals are copied verbatim, so that a comma or a word
			// that looks like a cell address inside one is left alone. A
			// doubled quote escapes a quote and does not end the literal.
			start := i
			i++
			for i < len(runes) {
				if runes[i] == '"' {
					if i+1 < len(runes) && runes[i+1] == '"' {
						i += 2
						continue
					}
					i++
					break
				}
				i++
			}
			tokens = append(tokens, token{tokenString, string(runes[start:i])})
		case isIdentStart(c):
			start := i
			for i < len(runes) && isIdentRune(runes[i]) {
				i++
			}
			tokens = append(tokens, token{tokenIdent, string(runes[start:i])})
		case c >= '0' && c <= '9':
			// Scanned separately from identifiers so that the decimal point
			// of a number is not mistaken for a sheet separator.
			start := i
			for i < len(runes) && (runes[i] == '.' || (runes[i] >= '0' && runes[i] <= '9')) {
				i++
			}
			tokens = append(tokens, token{tokenOther, string(runes[start:i])})
		default:
			tokens = append(tokens, token{tokenOther, string(c)})
			i++
		}
	}
	return tokens
}

// isIdentStart reports whether c can begin an identifier. "$" is included so
// that the absolute marker of "$A$1" stays part of the reference.
func isIdentStart(c rune) bool {
	return c == '_' || c == '$' || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c > 127
}

// isIdentRune reports whether c can continue an identifier. "." is included
// so that sheet-qualified references ("Sheet1.A1") stay a single token.
func isIdentRune(c rune) bool {
	return isIdentStart(c) || c == '.' || (c >= '0' && c <= '9')
}
