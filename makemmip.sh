#!/bin/sh

"/c/Program Files/7-Zip/7z.exe" a -tzip -- AutoPlayer.mmip $(find . -type f -name "*.vbs" -o -name "*.ini")