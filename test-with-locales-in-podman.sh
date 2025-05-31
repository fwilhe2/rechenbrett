#!/bin/bash

# SPDX-FileCopyrightText: 2025 Florian Wilhelm
#
# SPDX-License-Identifier: MIT

set -o errexit

# The purpose of this script is to avoid depending on the host's configuration of LibreOffice for running tests that depend on the locale

podman build -t test-rechenbrett .
podman run -it --rm -e LANG=en_US.UTF-8 localhost/test-rechenbrett:latest
podman run -it --rm -e LANG=de_DE.UTF-8 localhost/test-rechenbrett:latest
