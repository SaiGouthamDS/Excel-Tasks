package org.utility;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Base_Class {
	
	public static String date_Formate(Workbook wb, Cell cell, String value1) {
		DataFormat data = wb.createDataFormat();
		String format = data.getFormat(cell.getCellStyle().getDataFormat());
		
		switch (format) {
		case "m/d/yy":

			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/YYYY");
			 value1 = sdf.format(cell.getDateCellValue());
			break;

		case "mm/dd/yyyy":
			SimpleDateFormat sdf1 = new SimpleDateFormat("mm/dd/yyyy");
			value1 = sdf1.format(cell.getDateCellValue());
			break;

		case "dd/MM/yyyy":
			SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
			value1 = sdf2.format(cell.getDateCellValue());
			break;

		default:
			System.out.println("Date Format Not Match");
			break;
		}
		return value1;

	}
	
	
	
	
	public static String data_Formater(Cell cell) {
		
		DataFormatter df = new DataFormatter();
		String value1 = df.formatCellValue(cell);
		
		return value1;

	}
	
//	public static LinkedHashMap<String,String> get_String_Cell_Value(Row getRow,String columnName) throws IOException {
//		 LinkedHashMap<String,String> details = new LinkedHashMap<String,String>();
//		// String column_Name = PropertyFile.column_Name();
//		 
//		details.put(columnName, getRow.getCell(get_Column_Index()).getStringCellValue());
//		details.put("EEID", new DataFormatter().formatCellValue(getRow.getCell(0)));
//		details.put("Full name", getRow.getCell(1).getStringCellValue());
//		details.put("Job Title", getRow.getCell(2).getStringCellValue());
//		details.put("Department", getRow.getCell(3).getStringCellValue());
//		details.put("Business Unit", getRow.getCell(4).getStringCellValue());
//		details.put("Gender", getRow.getCell(5).getStringCellValue());
//		details.put("Ethnicity", getRow.getCell(6).getStringCellValue());
//		details.put("Age", new DataFormatter().formatCellValue(getRow.getCell(7)));
//
//		Cell cell = getRow.getCell(8);
//		String value = "";
//		String date_Formate = Base_Class.date_Formate(wb, cell, value);
//
//		details.put("Hire Day", date_Formate);
//		details.put("Annual Salary", new DataFormatter().formatCellValue(getRow.getCell(9)));
//		details.put("Bonus", new DataFormatter().formatCellValue(getRow.getCell(10)));
//		details.put("Country", getRow.getCell(11).getStringCellValue());
//		details.put("City", getRow.getCell(12).getStringCellValue());
//		details.put("Exit Date", new DataFormatter().formatCellValue(getRow.getCell(13)));
//		return details;
//
//
//	}
	
	public static void get_Numeric_Cell_Value(LinkedHashMap<String, String> details,Row getRow, String columnName, int cellNum) {
		details.put(columnName, new DataFormatter().formatCellValue(getRow.getCell(cellNum)));
		

	}
	
	public static void get_Date_Cell_value(LinkedHashMap<String, String> details,Row getRow,Workbook wb, String columnName, int cellNum) {
		Cell cell = getRow.getCell(cellNum);
		String value = "";
		String date_Formate = Base_Class.date_Formate(wb, cell, value);
		details.put(columnName, date_Formate);

	}
	
	
	
	
		

	

}
