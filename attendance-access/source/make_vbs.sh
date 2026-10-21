#!/bin/sh
# Converts the UTF-8 source into the UTF-16LE (with BOM, CRLF) script that
# Windows Script Host needs to read Arabic text correctly.
set -e
cd "$(dirname "$0")"
{ printf '\377\376'; sed 's/$/\r/' Build_Attendance_System.utf8.vbs | iconv -f UTF-8 -t UTF-16LE; } > ../Build_Attendance_System.vbs
