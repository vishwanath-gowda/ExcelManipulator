# ExcelManipulator


This is an API written over Apache POI to quickly access excel sheets without going much into the details of POI implementations.

As of now,

ExcelManipulator Object's 

ArraList\<Rows\> readAllRows(String sheetName,String excelPath);  will return an arrayList of Rows object, which in turn contains ArrayList of fields of one individual row.


boolean writeCell(String sheetName,String excelPath, int rowNo, int fieldNo,boolean createonNonExistence) function will write a particular cell in excel file and return tru on success else false.
