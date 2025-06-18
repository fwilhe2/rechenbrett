# SPDX-FileCopyrightText: 2025 Florian Wilhelm
#
# SPDX-License-Identifier: MIT

# The purpose of this container is to avoid depending on the host's configuration of LibreOffice for running tests that depend on the locale
# Based on the example golang Dockerfile on Docker Hub

FROM ghcr.io/fwilhe2/go-libreoffice:latest

WORKDIR /usr/src/app

# pre-copy/cache go.mod for pre-downloading dependencies and only redownloading them in subsequent builds if they change
COPY go.mod ./
RUN go mod download

COPY . .

ENTRYPOINT [ "go" ]
CMD [ "test", "-v", "./..." ]
