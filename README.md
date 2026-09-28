# Pivotmaker-for-Powerbi-v1
Version 1.0 of a python script that takes a .csv file exported from Powerbi website and produces a .xlxs file with a pivot table following the hierarchy of Dept Name > Cost Group > Cost Category and includes Grand Totals.

The script will process the first .xlsx file in the same folder that is labeled "data.xlsx" and export an .xlxs (Excel) file named "pivoted_report" followed by a timestamp. If no excel file under that name is found, it will search for a local .csv file with the same name.


Changelog:

Sept 28, 2026 - updated to accept .xlsx or .csv (before was only capable of .csv)

Sept 24, 2026 - initial commit
