# SPDX-FileCopyrightText: 2025 Florian Wilhelm
#
# SPDX-License-Identifier: MIT

all: format build test

format:
	gofumpt -w $$(find . -name '*.go')

build:
	go build -v ./...

test:
	go test -v ./...

showcase:
	go run ./cmd/showcase

clean:
	rm -rf _it-*