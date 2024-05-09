package org.utility;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

public class PropertyFile {
	private static final String config_Path ="C:\\Users\\sdevangasubramani\\eclipse-workspace\\Excel_Tasks\\Property_File\\Config.properties";
	
	
	
	public static Properties property_Path() throws IOException {
		Properties properties = new Properties();
		properties.load(Files.newInputStream(Paths.get(
				config_Path)));
		return properties;


	}
	
	public static String excelPath() throws IOException {
		return property_Path().getProperty("excel_Path");

	}
	
	public static String sheetName() throws IOException {
		return property_Path().getProperty("sheet");

	}
	
	
	
	
	public static String length_Of_Row_Start() throws IOException {
		return property_Path().getProperty("lengthOfRowStart");
		
	}
	
	public static String length_Of_Row_End() throws IOException {
		return property_Path().getProperty("lengthOfRowEnd");
		
	}
	
	public static Object emp_ID() throws IOException {
		return property_Path().get("empId");
	}
	
	public static Object detail_Of_Data() throws IOException {
		return property_Path().get("detail");
		

	}
	
	public static String column_Name() throws IOException {
		return property_Path().getProperty("columnName");

	}
	public static String uRL() throws IOException {
		return property_Path().getProperty("url");
	}
	
	public static int compareRow() throws IOException {
		String property = property_Path().getProperty("to_Compare_Particular_Row");
		int parseInt = Integer.parseInt(property);
		return parseInt;
		
	}
	
	public static String xPathForTheWebPage() throws IOException {
		return property_Path().getProperty("WebPage_Table_Row_Locator");

	}
	
	public static String startColumnName() throws IOException {
		return property_Path().getProperty("start_Column");

	}
	
	public static String endColumnName() throws IOException {
		return property_Path().getProperty("end_Column");

	}
	
	public static int startRowNumberToGetRangeValue() throws IOException {
		String property = property_Path().getProperty("length_Of_RowStart");
		int parseInt = Integer.parseInt(property);
		return parseInt;

	}
	
	
	public static int endRowNumberToGetRangeValue() throws IOException {
		String property = property_Path().getProperty("length_Of_RowEnd");
		int parseInt = Integer.parseInt(property);
		return parseInt;

	}
	
	

}
