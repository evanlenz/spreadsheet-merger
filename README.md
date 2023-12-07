# spreadsheet-merger
The run.sh script will merge two or more spreadsheets that you put in the
current directory as .xml files in Excel's "XML Spreadsheet 2003" format.
This is meant to support spreadsheets having column names that overlap with
each other and that can have variable numbers of repeating column headings
(using the same name).

## Practical application
My practical application was merging .csv files exported from Jira, which
only allows you to export 1000 records at a time, resulting in spreadsheets
that don't have the same position and number of column headings.

## Dependencies
You must have [SaxonJ-HE](https://www.saxonica.com/download/java.xml) installed.

You must have Saxon's jar file in your CLASSPATH environment variable. For example,
CLASSPATH=c:/saxon11/saxon-he-11.3.jar
