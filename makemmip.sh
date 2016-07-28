#!/bin/sh
version="0.1.0"

rm -f $(ls | grep "\.mmip$")
"/c/Program Files/7-Zip/7z.exe" a -tzip -- AutoPlayer-$version.mmip $(find . -type f -name "*.vbs" -o -name "*.ini")