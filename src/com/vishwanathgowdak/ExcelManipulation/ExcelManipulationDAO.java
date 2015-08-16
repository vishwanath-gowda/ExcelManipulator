package com.vishwanathgowdak.ExcelManipulation;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public interface ExcelManipulationDAO {
	
	public ArrayList<Rows> readAllRows(String sheetName, String excelPath) throws IOException;
	public int writeRow(String sheetName, String excelPath,ArrayList<String> values) throws IOException;
	public boolean writeCell(String sheetName,String excelPath,int rowNo,int fieldNo,String value, boolean createOnNonExistence) throws IOException;
	public ArrayList<Integer> getRowNo(String sheetName,String excelPath,int fieldNo,String value) throws IOException;
	public boolean deleteRow(String sheetName,String excelPath,int rowNo) throws IOException;
	public int getNoOfRows(String sheetName,String excelPath) throws  IOException;

}
