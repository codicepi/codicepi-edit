#!/bin/bash
pdftk $1 cat 1-200 output $1-first.pdf
pdftk $1 cat 201-end output $1-last.pdf

