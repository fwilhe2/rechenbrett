# SPDX-FileCopyrightText: 2025 Florian Wilhelm
#
# SPDX-License-Identifier: MIT

# The purpose of this container is to avoid depending on the host's configuration of LibreOffice for running tests that depend on the locale
# Based on the example golang Dockerfile on Docker Hub

FROM docker.io/golang:1.24

WORKDIR /usr/src/app

# pre-copy/cache go.mod for pre-downloading dependencies and only redownloading them in subsequent builds if they change
COPY go.mod ./
RUN go mod download

RUN apt-get -qq update && apt-get install --no-install-recommends -yqq libreoffice libreoffice-l10n-de

COPY . .

ENTRYPOINT [ "go" ]
CMD [ "test", "-v", "./..." ]
