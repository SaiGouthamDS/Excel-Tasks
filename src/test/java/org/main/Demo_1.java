package org.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
import org.testng.annotations.Test;
import org.utility.Base_Class;
import org.utility.Helper_Class;
import org.utility.PropertyFile;

public class Demo_1 {

	@Test
	private void entire_Excel_Data_Reader() throws IOException {

		Helper_Class.toGetAllDatasFromTheExcelSheet();

	}

	@Test
	public void to_Read_Particular_Rows() throws IOException {
		Helper_Class.to_Read_Particular_Rows_Values();

	}

	@Test
	private void to_Read_Particular_Column_Data_By_Employee_ID() throws IOException {
		String to_Read_Excel_Data = Helper_Class.to_Read_Excel_Data();
		System.out.println(to_Read_Excel_Data);

	}

	@Test
	public static void to_read_Excel_Particular_Row_Datas_By_Employee_ID() throws IOException {
		LinkedHashMap<String, String> to_Get_The_Employee_Details = Helper_Class.to_Get_The_Employee_Details();
		Set<Entry<String, String>> entrySet = to_Get_The_Employee_Details.entrySet();
		for (Entry<String, String> entry : entrySet) {
			System.out.println(entry.getKey() + " : " + entry.getValue());

		}

	}
	
	
	@Test
	public static void toCompareExcelDataWithWebPageData() throws IOException {
		Helper_Class.to_Compare_Particular_Excel_Row_Datas_With_The_Webpage_Datas();

	}
	
	@Test
	private void getTheRangeValuesFromTheExcelSheet() throws IOException {
		Helper_Class.toGetRangeDetailsFromTheExcelSheet();

	}
	
	
	
	
	
	
	
	
	
	
	
	@Test
	private void creatNewExcelSheet() throws IOException {
	Helper_Class.toCreateNewExcelSheet();

	}
	
	@Test
	private void CreateAndWriteNewCellIntoExcelSheetWithExistingRow() throws IOException {
		Helper_Class.toCreateAndWriteIntoNewExcelSheetWithExistingRow();

	}
	
	@Test
	private void creatNewRowWithFirstCell() throws IOException {
		Helper_Class.toCreateNewRowAndFirstCellIntoExcelSheet();

	}
	
	@Test
	private void updateTheCellValue() throws IOException {
		Helper_Class.toUpdateTheValueInExcelSheet();
		
	}
	
	
	@Test
	private void test_Case_to_Window() {
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		driver.get("https://practice.expandtesting.com/windows");
		driver.manage().window().maximize();
		String parent_title = driver.getTitle();
		System.out.println(parent_title);
		WebElement clkhere = driver.findElement(By.xpath("//a[text()='Click Here']"));
		clkhere.click();
		String currentWindowHandle = driver.getWindowHandle();
		Set<String> allWindowHandles = driver.getWindowHandles();
		int size = allWindowHandles.size();
		System.out.println(size);
		for (String windowHandle : allWindowHandles) {
		    if (!windowHandle.equals(currentWindowHandle)) {
		        driver.switchTo().window(windowHandle);
		        // Perform actions in the new window
		    }
		}
		String window_title = driver.getTitle();
		System.out.println(window_title);
		WebElement exm_new_window = driver.findElement(By.xpath("//h1[text()='Example of a new window']"));
		String text = exm_new_window.getText();
		System.out.println(text);

		driver.quit();
	}
	
	
	@Test
	private void toCheckThePageTitle() {
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		driver.get("https://www.amazon.in/");
		driver.manage().window().maximize();
		String parent_title = driver.getTitle();
		System.out.println(parent_title);
		WebElement text_in_serach = driver.findElement(By.xpath("//input[@aria-label='Search Amazon.in']"));
		text_in_serach.sendKeys("Iphone");
		WebElement clk_search  = driver.findElement(By.xpath("//input[@id='nav-search-submit-button']"));
		List<WebElement> all_products = driver.findElements(By.xpath("//span[@class='a-size-medium a-color-base a-text-normal']"));
		System.out.println(all_products.size());
		
		driver.quit();
		

	}
	
}
	
	
	
	
	
	















































	
	
	
	
	
	
	
	
	
	
//	
//	@Test
//	private void demo() throws IOException {
//		
//		File f = new File("C:\\Users\\sdevangasubramani\\eclipse-workspace\\Excel_Tasks\\Excel\\DemoDataSheet.xlsx");
//		FileInputStream fis = new FileInputStream(f);
//		Workbook wb = new XSSFWorkbook(fis);
//		Sheet sheet = wb.getSheet("Data Sheet");
//		int lastRowNum = sheet.getLastRowNum();
//		short lastCellNum = sheet.getRow(0).getLastCellNum();
//		int column_range_start_Index = 0;
//		int column_range_End_Index = 0;
//		for (int i = 0; i < lastCellNum; i++) {
//			Cell cell = sheet.getRow(0).getCell(i);
//			
//			String stringCellValue = cell.getStringCellValue();
//			
//			if (stringCellValue.equalsIgnoreCase("Full Name")) {
//				
//				column_range_start_Index = cell.getColumnIndex();
//				
//				
//			}
//			else if (stringCellValue.equalsIgnoreCase("City")) {
//				column_range_End_Index = cell.getColumnIndex();
//				
//				
//				
//			}
//			
//			System.out.println(column_range_start_Index);
//			System.out.println(column_range_End_Index);
//			
//			
//			
//			
//			
//		}
//		System.out.println(column_range_start_Index);
//		System.out.println(column_range_End_Index);
//		
//		
//
//		System.out.println("===============================================");
//		}
//		fis.close();
//		System.out.println("Current Number Of Rows -" + lastRowNum);
//		System.out.println("Current Number Of Cells-" + lastCellNum);
//		System.out.println("===============================================");
//	
//	}
//	
//
//	
//	
//	
//	
	
	
	
		
//			StringBuffer sb = new StringBuffer();
//			//System.out.println(s[i]);
//			Boolean flag = Character.isDigit(s[1].charAt(0));
//			//System.out.println(flag);
//			if (flag) {
//				
//				SimpleDateFormat sm = new SimpleDateFormat("dd/MM/yyyy");
//				String format = sm.format(cell2.getDateCellValue());
//				System.out.println(format);
//				
//				
//				
//				
//			}
//			else {
//				SimpleDateFormat sm = new SimpleDateFormat("MM/dd/yyyy");
//				String format = sm.format(cell2.getDateCellValue());
//				System.out.println(format);
//			}
			
		

			
			
			
			
			
		
		
		
		
		
		

	
	
	
	
	

