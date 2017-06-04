#!/bin/sh
version=$(git describe --dirty)

rm $(ls . | grep "\.mmip$")
"/c/Program Files/7-Zip/7z.exe" a -tzip -- AutoPlayer-$version.mmip $(find . -type f -name "*.vbs" -o -name "*.ini") "changelog.md"
