## Overview
CPMGRID is a demo version of a 64 bit Excel add-in with predefined queries for a P6 Professional standalone <br> SQLite database. Queries provided 
allow users to get information from a single schedule or compare two different schedules. 

This add-in is built on SQLite for Excel from Govert van Drimmelen. <br>
https://github.com/govert/SQLiteForExcel

The add-in also uses XLSQLite from Mark Camilleri.

## Why is another comparison tool needed?
There are a number of programs on the market that can compare two P6 schedules including P6 itself. The output <br>
format of the file from P6 is not the greatest. Other programs can output data to an Excel file which works fine, <br>
but users inevitably end up performing analysis in Excel. One goal of CPMGRID is to start in Excel and stay <br>
where the data is going to be analyzed. 

## Installation
Three DLL files need to be placed in the same location as the xlam file.
* The Microsoft .NETFramework 3.5 needs to be installed
* SQLite3_StdCall.dll - from the SQLite for Excel library located at: <br> https://github.com/govert/SQLiteForExcel
* sqlite3.dll - This is the latest 64-bit DLL for SQLite version 3.42.0 located at: <br> https://www.sqlite.org/download.html
* cpmgrid_xlam_64.dll - This is the compiled VBA code


## Available Single Project Queries
* Negative total float
* Original Duration over 20 working days
* Original durations equal to remaining duration
* Out of sequence activities
* Remaining durations greater than original duration
* Duplicate activity names
* Level of effort activities
* Milestones
* Constraints
* Dangling starts and finishes

## Available Two Schedule Comparisons
* Renamed activities
* Added activities
* Deleted activities
* Original duration
* Remaining duration
* Added successors
* Deleted successors
* Activity budgeted cost
* Lags
* Actual starts and finishes
* Unchanged remaining duration
* Unchanged physical percent complete
* Calendar Names
* Missed late starts and finishes