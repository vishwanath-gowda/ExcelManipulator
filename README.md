# ExcelManipulator


This is an API written over Apache POI to quickly access excel sheets without going much into the details of POI implementations.

As of now,

Following are the methods provided by ExcelManipulation class.


public ArrayList\<Rows\> readAllRows(String sheetName, String excelPath) throws IOException;

                  This method return an arrayList of Rows object, Rows object contains a ArrayList of field containing each and every cell value of a row
  

public boolean writeCell(String sheetName,String excelPath,int rowNo,int fieldNo,String value, boolean createOnNonExistence) throws IOException;

                Writes String data to the specified cell. returns true on success else false.


public ArrayList\<Integer\> getRowNo(String sheetName,String excelPath,int fieldNo,String value) throws IOException;
                  This method returns the ArrayList containing the row numbers of all rows which contains "value" value in cellNo "field No"
