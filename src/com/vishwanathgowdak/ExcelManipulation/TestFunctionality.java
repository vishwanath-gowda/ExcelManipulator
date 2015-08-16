package com.vishwanathgowdak.ExcelManipulation;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Iterator;

public class TestFunctionality {

	public static void main(String[] args) {
		ExcelManipulation excelManipulation=new ExcelManipulation();
		
		try {
			//System.out.println(excelManipulation.writeCell("a", "d:\\abc.xlsx", 0, 0, "0",true));
//			Iterator<Rows> rows=excelManipulation.readAllRows("a", "d:\\abc.xlsx").iterator();
//			while(rows.hasNext()){
//				Rows row=rows.next();
//				Iterator<String> cellIterator=row.fields.iterator();
//				while (cellIterator.hasNext()){
//
//					System.out.println(cellIterator.next());
//				}
//			}
			System.out.println(excelManipulation.getNoOfRows("a", "d:\\abc.xlsx"));
			System.out.println(excelManipulation.deleteRow("a", "d:\\abc.xlsx", 1));
			System.out.println(excelManipulation.getNoOfRows("a", "d:\\abc.xlsx"));
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
