#!/bin/bash

java net.sf.saxon.Transform -s:merge-spreadsheets.xsl -xsl:merge-spreadsheets.xsl -o:output/merged.xml
