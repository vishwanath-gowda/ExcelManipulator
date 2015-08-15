package com.vishwanathgowdak.ExcelManipulation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelManipulation implements ExcelManipulationDAO{


	private XSSFWorkbook workbook=null;
	private XSSFSheet sheet =null;
	private Iterator<Row> rowIterator =null;
	private Iterator<Cell> cellIterator=null;
	Rows row=null;
	ArrayList<Rows> temp=null;
	@Override
	public ArrayList<Rows> readAllRows(String sheetName, String excelPath) throws IOException {
		FileInputStream file = new FileInputStream(new File(excelPath));
		try{
			temp= new ArrayList<Rows>();
			workbook= new XSSFWorkbook(file);
			sheet = workbook.getSheet(sheetName);
			rowIterator = sheet.iterator();

			while(rowIterator.hasNext()){
				row = new Rows();
				cellIterator = rowIterator.next().cellIterator();
				while(cellIterator.hasNext()){
					row.fields.add(cellIterator.next().toString());

				}
				temp.add(row);
				row=null;
			}
		}catch(IOException e){
			throw e;
		}finally{
			if(workbook!=null)
				workbook.close();
			if(file!=null)
				file.close();

		}
		return temp;

	}

	@Override
	public int writeRow(String sheetName, String excelPath,ArrayList<String> values)throws IOException {
		
		return 0;
	}

	@Override
	public boolean writeCell(String sheetName, String excelPath, int rowNo,
			int fieldNo, String value, boolean createOnNonExistence) throws IOException {
		FileInputStream file=null;
		XSSFWorkbook wb=null;;
		XSSFSheet sheet;

		try{
			file = new FileInputStream(new File(excelPath));
			wb= new XSSFWorkbook(file);
			sheet = wb.getSheet(sheetName);
			if(sheet.getRow(rowNo)==null && createOnNonExistence){
				sheet.createRow(rowNo);
			}else{
				return false;
			}
			if(sheet.getRow(rowNo).getCell(fieldNo)==null&& createOnNonExistence){
				sheet.getRow(rowNo).createCell(fieldNo);
			}else{
				return false;
			}

			Cell cell= sheet.getRow(rowNo).getCell(fieldNo);
			cell.setCellValue(value);
			file.close();
			FileOutputStream outFile =new FileOutputStream(new File(excelPath));
			wb.write(outFile);
			outFile.close();
		}catch(Exception e){
			throw e;
		}finally{
			if(wb!=null)
				wb.close();	
			if(file!=null)
				file.close();
		}

		return true;
	}

}