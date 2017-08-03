# SpreadSheet (Excel 2007 and above) Importer
This is a simple Importer (load data from spreadsheet to target database) implementation with EPPlus library. There're two major components:

- DbMapper - define the mapping rule between spreadsheet and database in DbMapper.xml
- Importer - loading spreadsheet and parsing the data in excel and then insert data into database with injected DbMapper

**NOTE: need to restore the package for getting expected EPPlus lib**