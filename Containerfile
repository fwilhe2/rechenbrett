FROM docker.io/golang:1.24

WORKDIR /usr/src/app

# pre-copy/cache go.mod for pre-downloading dependencies and only redownloading them in subsequent builds if they change
COPY go.mod ./
RUN go mod download

RUN apt-get -qq update && apt-get install --no-install-recommends -yqq libreoffice libreoffice-l10n-de

COPY . .

ENTRYPOINT [ "go" ]
CMD [ "test", "-v", "./..." ]
