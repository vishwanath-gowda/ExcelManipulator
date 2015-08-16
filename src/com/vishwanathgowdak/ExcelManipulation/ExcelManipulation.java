package com.vishwanathgowdak.ExcelManipulation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelManipulation implements ExcelManipulationDAO{




	@Override
	public ArrayList<Rows> readAllRows(String sheetName, String excelPath) throws IOException {
		ArrayList<Rows> temp=null;
		XSSFWorkbook workbook=null;
		XSSFSheet sheet =null;
		Iterator<Row> rowIterator =null;
		Iterator<Cell> cellIterator=null;
		Rows row=null;
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
			if(sheet.getRow(rowNo)==null && !createOnNonExistence){
				return false;
			}else if(sheet.getRow(rowNo)==null && createOnNonExistence){
				sheet.createRow(rowNo);
			}
			if(sheet.getRow(rowNo).getCell(fieldNo)==null && !createOnNonExistence){
				return false;
			}else if(sheet.getRow(rowNo).getCell(fieldNo)==null && createOnNonExistence){
				sheet.getRow(rowNo).createCell(fieldNo);
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

	@Override
	public ArrayList<Integer> getRowNo(String sheetName, String excelPath, int fieldNo,String value) throws IOException {

		XSSFWorkbook wb=null;
		XSSFSheet s=null;
		ArrayList<Integer> temp=new ArrayList<Integer>();
		File excel=new File(excelPath);
		if(!excel.exists()){
			throw new FileNotFoundException();
		}
		FileInputStream file = new FileInputStream(excel);
		try{
			wb = new XSSFWorkbook(file);
			s = wb.getSheet(sheetName);
			Iterator<Row> rowI = s.iterator();
			while(rowI.hasNext()){
				Row row=rowI.next();
				if(row.getCell(fieldNo).getStringCellValue()==value)
					temp.add(row.getRowNum());
			}
			return temp;
		}catch(Exception e){
			throw e;
		}finally{
			wb.close();
		}

	}

	@Override
	public boolean deleteRow(String sheetName, String excelPath, int rowNo)
			throws IOException {

		XSSFWorkbook workbook=null;
		XSSFSheet sheet =null;
		FileInputStream file = new FileInputStream(new File(excelPath));
		try{
			file = new FileInputStream(new File(excelPath));
			workbook= new XSSFWorkbook(file);
			sheet = workbook.getSheet(sheetName);
			if(sheet==null){
				System.out.println("returning");
				return false;
			}
			int lastRowNum=sheet.getLastRowNum();
			if(rowNo>=0&&rowNo<lastRowNum){
				sheet.shiftRows(rowNo+1,lastRowNum, -1);
			}
			if(rowNo==lastRowNum){
				XSSFRow removingRow=sheet.getRow(rowNo);
				if(removingRow!=null){
					sheet.removeRow(removingRow);
				}
			}
			file.close();
			FileOutputStream outFile =new FileOutputStream(new File(excelPath));
			workbook.write(outFile);
			outFile.close();
			return true;

		}catch(Exception e){
			throw e;
		}finally{
			if(workbook!=null)
				workbook.close();
		}

	}

	@Override
	public int getNoOfRows(String sheetName, String excelPath) throws FileNotFoundException, IOException {
		XSSFWorkbook wb=null;
		try{
			wb= new XSSFWorkbook( new FileInputStream(new File(excelPath)));
			int rowCount=wb.getSheet(sheetName).getLastRowNum() +1;
			return rowCount;
		}finally{
			wb.close();
		}


	}

}
