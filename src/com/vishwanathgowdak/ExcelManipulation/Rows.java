package com.vishwanathgowdak.ExcelManipulation;

import java.util.ArrayList;

public class Rows {

 ArrayList<String> fields= new ArrayList<String>();
 public Rows(){}
 public Rows(ArrayList<String> fields){
	 this.fields=fields;
 }

public ArrayList<String> getFields() {
	return fields;
}

public void setFields(ArrayList<String> fields) {
	this.fields = fields;
}
 
 
}
