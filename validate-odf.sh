#!/bin/bash

# SPDX-FileCopyrightText: 2025 Florian Wilhelm
#
# SPDX-License-Identifier: MIT

set -o errexit

podman run -t --mount=type=bind,source=$PWD,destination=/usr/src/data ghcr.io/fwilhe2/odfvalidator-action:latest
