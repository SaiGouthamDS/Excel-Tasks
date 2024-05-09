package org.utility;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

public class ToCreateNewExcelSheet {
	
	private static final String config_Path_Create = "C:\\Users\\sdevangasubramani\\eclipse-workspace\\Excel_Tasks\\Property_File\\tocreat.properties";
	
	public static Properties propertyPathforNewcreateExcel() throws IOException {
		Properties properties = new Properties();
		properties.load(Files.newInputStream(Paths.get(
				config_Path_Create)));
		return properties;

	}
	
	public static String creatExcel() throws IOException {
		
		return propertyPathforNewcreateExcel().getProperty("create_Excel_Path");
		

	}
	
	public static String creatSheet() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("create_Sheet");

	}
	
	public static int creatRow() throws IOException {
		String propertyRow = propertyPathforNewcreateExcel().getProperty("create_Row");
		int row_num = Integer.parseInt(propertyRow);
		return row_num;

	}
	
	public static int creatCell() throws IOException {
		String propertyCell = propertyPathforNewcreateExcel().getProperty("create_Cell");
		int cell_num = Integer.parseInt(propertyCell);
		return cell_num;
	}
	
	public static String writeValueOnExcel() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("write_value");
		

	}
	
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	public static String getExcelPathForCellCreation() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("excel_Path_for_Cell_Creation");
		

	}
	
	
	public static String getSheetNameForCellcreation() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("sheet_Name_for_Cell_Creation");
	}
	
	public static int getRowNum() throws IOException {
		String propertyRow = propertyPathforNewcreateExcel().getProperty("row_Num");
		int row_num = Integer.parseInt(propertyRow);
		return row_num;
		

	}
	
	public static int creatCellNum() throws IOException {
		String propertyCell = propertyPathforNewcreateExcel().getProperty("create_Cell_Num");
		int cell_num = Integer.parseInt(propertyCell);
		return cell_num;
		
	}
	
	public static String writeValueInNewCell() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("write_value_in_New_Cell");

	}
	
	//////////////////////////////// Create a New Row with first Cell
	
	
	public static String excelPathforNewRowCreation() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("excel_Path_for_New_Row_Creation");

	}
	
	public static String sheetNameForNewRowCreation() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("sheet_Name_For_New_Row_Creation");

	}
	
	public static int createNewRow() throws IOException {
		String propertyRow = propertyPathforNewcreateExcel().getProperty("create_New_Row");
		int row_num = Integer.parseInt(propertyRow);
		return row_num;
		
		

	}
	
	public static int createNewCell() throws IOException {
		String propertyCell = propertyPathforNewcreateExcel().getProperty("create_New_Cell");
		int cell_num = Integer.parseInt(propertyCell);
		return cell_num;
	}
	
	public static String writevalueinnewRowandFirstCell() throws IOException {
	return propertyPathforNewcreateExcel().getProperty("write_value_in_new_Row_and_First_Cell");
	}
	
	
	////////////////Update The Cell Value/////////
	
	public static String excelPathForUpadetValue() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("excel_Path_For_Update_Cell_Value");

	}
	
	public static String sheetForUpdateValue() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("sheet_Name_For_Update_Cell_Value");

	}
	
	public static int updateNoOfRow() throws IOException {
		String propertyRow = propertyPathforNewcreateExcel().getProperty("update_No_Of_Row");
		int row_num = Integer.parseInt(propertyRow);
		return row_num;
	}
	
	public static int updateNoOfCell() throws IOException {
		
		String propertyCell = propertyPathforNewcreateExcel().getProperty("update_No_Of_Cell");
		int cell_num = Integer.parseInt(propertyCell);
		return cell_num;

	}
	
	public static String updateTheNewValue() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("new_Value_For_Update");

	}
	
	public static String oldData() throws IOException {
		return propertyPathforNewcreateExcel().getProperty("old_Cell_data");

	}
	
	
	
	

}
