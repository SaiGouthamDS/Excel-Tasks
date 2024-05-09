package org.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.reporters.jq.Main;

public class Helper_Class extends PropertyFile {

	public static void to_read_Data_from_Excel() throws IOException {

		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		int lastRowNum = sheet.getLastRowNum();
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		for (int i = 0; i < lastRowNum; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < lastCellNum; j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				String value = "";
				if (cellType == 1) {
					DataFormatter d = new DataFormatter();
					value = d.formatCellValue(cell);
				}
				else if (DateUtil.isCellDateFormatted(cell)) {
					value = Base_Class.date_Formate(wb, cell, value);
				}
				else {
					value = Base_Class.data_Formater(cell);
				}
		System.out.println(sheet.getRow(0).getCell(j)+" : "+value);
			}
		System.out.println("===============================================");
		}
		fis.close();
		System.out.println("Current Number Of Rows -" + lastRowNum);
		System.out.println("Current Number Of Cells-" + lastCellNum);
		System.out.println("===============================================");
	}
	
	
	
	
	
	public static void to_Read_Particular_Rows_Values() throws IOException {
		
		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		int lastRowNum = sheet.getLastRowNum();
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		for (int i = Integer.parseInt(length_Of_Row_Start()); i < Integer.parseInt(length_Of_Row_End()); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < lastCellNum-1; j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				String value = "";
				if (cellType == 1) {
					value = cell.getStringCellValue();
				}
				else if (DateUtil.isCellDateFormatted(cell)) {
					DataFormatter d = new DataFormatter();
					value = d.formatCellValue(cell);
				}
				else {
					value = Base_Class.data_Formater(cell);
				}
		System.out.println(sheet.getRow(0).getCell(j)+" : "+value);
			}
			System.out.println("===============================================");
		}
		fis.close();
	}

	
	
	
	public static String to_Read_Excel_Data() throws IOException {

		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		LinkedHashMap<String, LinkedHashMap<String, String>> empid = new LinkedHashMap<String, LinkedHashMap<String, String>>();
		String column_Name = column_Name();
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			Row getRow = sheet.getRow(i);
			LinkedHashMap<String, String> details = new LinkedHashMap<String, String>();
			for (int j = 0; j < lastCellNum-1; j++) {
				Cell cell = sheet.getRow(0).getCell(j);
				String string_col = cell.getStringCellValue();
				if (string_col.equalsIgnoreCase(column_Name)) {
					int columnIndex = cell.getColumnIndex();
					Cell cell2 = getRow.getCell(columnIndex);
					int cellType = cell2.getCellType();
					if (cellType == 1) {
						details.put(column_Name, getRow.getCell(columnIndex).getStringCellValue());
					}
					else {
						details.put(column_Name, new DataFormatter().formatCellValue(getRow.getCell(columnIndex)));
					}
				}
			}
			empid.put(String.valueOf(sheet.getRow(i).getCell(0)), details);
		}
		fis.close();
		return empid.get(emp_ID()).get(column_Name);
	}

	
	
	
	public static LinkedHashMap<String, String> to_Get_The_Employee_Details() throws IOException {
		
		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		LinkedHashMap<String, LinkedHashMap<String, String>> empid = new LinkedHashMap<String, LinkedHashMap<String, String>>();
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			Row getRow = sheet.getRow(i);
			LinkedHashMap<String, String> details = new LinkedHashMap<String, String>();
			for (int j = 0; j < lastCellNum-1; j++) {
				Cell cell = sheet.getRow(0).getCell(j);
				String stringCellValue = cell.getStringCellValue();
				int columnIndex = cell.getColumnIndex();
				Cell cell2 = getRow.getCell(columnIndex);
				int cellType = cell2.getCellType();
					if (cellType == 1) {
						details.put(stringCellValue, getRow.getCell(columnIndex).getStringCellValue());
					}
					else {
						details.put(stringCellValue, new DataFormatter().formatCellValue(getRow.getCell(columnIndex)));
					}	
				}
			empid.put(String.valueOf(sheet.getRow(i).getCell(0)), details);
		}
		fis.close();
		return empid.get(emp_ID());
	}
	
	
	
	
	
	public static void to_Compare_Particular_Excel_Row_Datas_With_The_Webpage_Datas() throws IOException {

		// Excel Data
		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());

		// WebPage Data
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get(uRL());
		driver.manage().window().maximize();
		List<WebElement> webpage_table_datas = driver.findElements(By.xpath(xPathForTheWebPage()));
		for (int i = 0; i < webpage_table_datas.size(); i++) {
			Row row = sheet.getRow(compareRow());
			String text = webpage_table_datas.get(i).getText();
			Cell cell = row.getCell(i);
			// Below codes are used to convert the string values,
			// if present Date & Numeric Cell datas on the Excel Sheet
			int cellType = cell.getCellType();
			String value = "";
			if (cellType == 1) {
				value = cell.getStringCellValue();
			}
			else if (DateUtil.isCellDateFormatted(cell)) {
				DataFormatter d = new DataFormatter();
				value = d.formatCellValue(cell);
			}
			else {
				double ncv = cell.getNumericCellValue();
				long l = (long) ncv;
				value = String.valueOf(l);
			}
			if (text.equalsIgnoreCase(value)) {
				System.out.println("Data is Matched------" + "(Excel Data---" + value + ")" + " & " + "(WebPage Data---"
						+ text + ")" + ",Cell Number - " + i);
			} else {
				System.out.println("Data Does not Match-----" + "(Excel Data---" + value + ")" + " & "
						+ "(WebPage Data---" + text + ")" + ",Cell Number - " + i);
			}
		}
		driver.close();
	}

	
	public static void toGetAllDatasFromTheExcelSheet() throws IOException {
		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		int lastRowNum = sheet.getLastRowNum();
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j <lastCellNum; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					System.out.println(sheet.getRow(0).getCell(j)+" : "+cell+"(Cell is Empty)");
				} else {
					String value="";
					int cellType = cell.getCellType();
					if (cellType == 1) {
						DataFormatter d = new DataFormatter();
						value = d.formatCellValue(cell);
					}
					else if (DateUtil.isCellDateFormatted(cell)) {
						value = Base_Class.date_Formate(wb, cell, value);
					}
					else {
						value = Base_Class.data_Formater(cell);
					}
					System.out.println(sheet.getRow(0).getCell(j)+" : "+value);
				}
			
			}
			
		System.out.println("===============================================");
		}
		fis.close();
		System.out.println("Current Number Of Rows -" + lastRowNum);
		System.out.println("Current Number Of Cells-" + lastCellNum);
		System.out.println("===============================================");
	}
	
	
	
	
	public static void toGetRangeDetailsFromTheExcelSheet() throws IOException {
		File f = new File(excelPath());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		int lastRowNum = sheet.getLastRowNum();
		short lastCellNum = sheet.getRow(0).getLastCellNum();
		int column_range_start_Index=0;
		int column_range_End_Index=0;
		for (int i = 0; i < lastCellNum; i++) {
			Cell cell = sheet.getRow(0).getCell(i);
			String stringCellValue = cell.getStringCellValue();
			if (stringCellValue.equalsIgnoreCase(startColumnName())) {
				column_range_start_Index = cell.getColumnIndex();
			} 
			else if (stringCellValue.equalsIgnoreCase(endColumnName())) {
				column_range_End_Index = cell.getColumnIndex();
			}
		}
		for (int i = startRowNumberToGetRangeValue(); i <=endRowNumberToGetRangeValue(); i++) {
			Row row = sheet.getRow(i);
			for (int j = column_range_start_Index; j <=column_range_End_Index; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					System.out.println(sheet.getRow(0).getCell(j)+" : "+cell+"(Cell is Empty)");
				} else {
					String value="";
					int cellType = cell.getCellType();
					
					if (cellType == 1) {
						DataFormatter d = new DataFormatter();
						value = d.formatCellValue(cell);
					}
					else if (DateUtil.isCellDateFormatted(cell)) {
						value = Base_Class.date_Formate(wb, cell, value);
					}
					else {
						value = Base_Class.data_Formater(cell);
					}
					System.out.println(sheet.getRow(0).getCell(j)+" : "+value);
				}
			}
		System.out.println("===============================================");
		}
		fis.close();
		System.out.println("Current Number Of Rows -" + lastRowNum);
		System.out.println("Current Number Of Cells-" + lastCellNum);
		System.out.println("===============================================");
	}
	
	
	

	
	public static void toCreateNewExcelSheet() throws IOException {
		File f = new File(ToCreateNewExcelSheet.creatExcel());
		Workbook wb = new XSSFWorkbook();
		Sheet createSheet = wb.createSheet(ToCreateNewExcelSheet.creatSheet());
		Row createRow = createSheet.createRow(ToCreateNewExcelSheet.creatRow());
		Cell createCell = createRow.createCell(ToCreateNewExcelSheet.creatCell());
		createCell.setCellValue(ToCreateNewExcelSheet.writeValueOnExcel());
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		System.out.println("New Excel File Created");
	}
	
	
	
	
	public static void toCreateAndWriteIntoNewExcelSheetWithExistingRow() throws IOException {
		File f = new File(ToCreateNewExcelSheet.getExcelPathForCellCreation());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet get_sheet = wb.getSheet(ToCreateNewExcelSheet.getSheetNameForCellcreation());
		Row get_row = get_sheet.getRow(ToCreateNewExcelSheet.getRowNum());
		Cell creat_Cell = get_row.createCell(ToCreateNewExcelSheet.creatCellNum());
		creat_Cell.setCellValue(ToCreateNewExcelSheet.writeValueInNewCell());
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		System.out.println("Create and Write a new Cell Into The Excel Sheet on already Existing Row");
	}
	
	
	
	
	public static void toCreateNewRowAndFirstCellIntoExcelSheet() throws IOException {
		File f = new File(ToCreateNewExcelSheet.excelPathforNewRowCreation());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet get_sheet = wb.getSheet(ToCreateNewExcelSheet.sheetNameForNewRowCreation());
		Row creat_Row = get_sheet.createRow(ToCreateNewExcelSheet.createNewRow());
		Cell create_Cell = creat_Row.createCell(ToCreateNewExcelSheet.createNewCell());
		create_Cell.setCellValue(ToCreateNewExcelSheet.writevalueinnewRowandFirstCell());
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		System.out.println("Created the New Row");
	}
	
	
	
	
	
	public static void toUpdateTheValueInExcelSheet() throws IOException {
		File f = new File(ToCreateNewExcelSheet.excelPathForUpadetValue());
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet get_sheet = wb.getSheet(ToCreateNewExcelSheet.sheetForUpdateValue());
		int lastRowNum = get_sheet.getLastRowNum();
		short lastCellNum = get_sheet.getRow(0).getLastCellNum();
		for (int i = 0; i < lastRowNum; i++) {
			Row row = get_sheet.getRow(i);
			for (int j = 0; j < lastCellNum; j++) {
				Cell cell = row.getCell(j);
				String cell_value = cell.getStringCellValue();
				if (cell_value.equalsIgnoreCase(ToCreateNewExcelSheet.oldData())) {
					cell.setCellValue(ToCreateNewExcelSheet.updateTheNewValue());
				}
			}
			
		}	
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		
		System.out.println("Updated The Value");
		
		
	}
	
	
	
	

	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	public static LinkedHashMap<String, String> to_Read_Excel_Row_Datas() throws IOException {
		

		
		File f = new File(excelPath());

		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet(sheetName());
		
		LinkedHashMap<String,LinkedHashMap<String,String>> empid = new LinkedHashMap<String,LinkedHashMap<String,String>>();
		

		for (int i = 1; i < sheet.getLastRowNum(); i++) {

			Row getRow = sheet.getRow(i);

			LinkedHashMap<String, String> details = new LinkedHashMap<String, String>();
			
			
			//Base_Class.get_String_Cell_Value(details, getRow, column_Name(), Base_Class.get_Column_Index());
			
			
			
			
//			details.put("EEID", new DataFormatter().formatCellValue(getRow.getCell(0)));
//			details.put("Full name", getRow.getCell(1).getStringCellValue());
//			details.put("Job Title", getRow.getCell(2).getStringCellValue());
//			details.put("Department", getRow.getCell(3).getStringCellValue());
//			details.put("Business Unit", getRow.getCell(4).getStringCellValue());
//			details.put("Gender", getRow.getCell(5).getStringCellValue());
//			details.put("Ethnicity", getRow.getCell(6).getStringCellValue());
//			details.put("Age", new DataFormatter().formatCellValue(getRow.getCell(7)));
//
//			Cell cell = getRow.getCell(8);
//			String value = "";
//			String date_Formate = Base_Class.date_Formate(wb, cell, value);
//
//			details.put("Hire Day", date_Formate);
//			details.put("Annual Salary", new DataFormatter().formatCellValue(getRow.getCell(9)));
//			details.put("Bonus", new DataFormatter().formatCellValue(getRow.getCell(10)));
//			details.put("Country", getRow.getCell(11).getStringCellValue());
//			details.put("City", getRow.getCell(12).getStringCellValue());
//			details.put("Exit Date", new DataFormatter().formatCellValue(getRow.getCell(13)));

			empid.put(String.valueOf(sheet.getRow(i).getCell(0)), details);

		}
		fis.close();

		return empid.get(emp_ID());
	}
	
	
	

}
