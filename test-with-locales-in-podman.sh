#!/bin/bash

set -o errexit

podman build -t test-rechenbrett .
podman run -it --rm -e LANG=en_US.UTF-8 localhost/test-rechenbrett:latest
podman run -it --rm -e LANG=de_DE.UTF-8 localhost/test-rechenbrett:latest
