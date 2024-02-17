package exceldataprovider;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import org.apache.poi.ss.usermodel.CellType;

import io.github.bonigarcia.wdm.WebDriverManager;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelDataProviderDemo {
	
	@Test(dataProvider = "getData")
	public void loginTest(String email,Object password)
	{
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.get("https://tutorialsninja.com/demo/");
		
		driver.findElement(By.xpath("//span[text()= 'My Account']")).click();
		driver.findElement(By.linkText("Login")).click();
		driver.findElement(By.id("input-email")).sendKeys(email);
		driver.findElement(By.id("input-password")).sendKeys(password.toString());
		driver.findElement(By.xpath("//input[@value = 'Login']")).click();
		
		Assert.assertTrue(driver.findElement(By.linkText("Edit your account information")).isDisplayed());
		driver.quit();
	}
	
	@DataProvider(name ="getData",parallel=true)
	public Object[][] getData()
	{
		String excelPath= System.getProperty("user.dir")+"\\src\\test\\resources\\tutorialsninjadata.xlsx";
	
		File file = new File(excelPath);
		FileInputStream fis = null;
		try {
			 fis = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		XSSFWorkbook workbook =null;
		try {
			  workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount= sheet.getRow(0).getLastCellNum();
		
		Object[][] obj= new Object[rowCount-1][colCount];
		
		Iterator<Row> rows = sheet.iterator();
		
		int r=0;
		int c=0;
		
		while(rows.hasNext())
		{
			Row row = rows.next();
			
			if(r==0)
			{
				if(rows.hasNext())
				{
					row=rows.next();
				}
				
			}
			
			Iterator<Cell> cells =row.iterator();
			while(cells.hasNext())
			{
				Cell cell = cells.next();
				
				switch(cell.getCellType())
				{
					case STRING : 
					obj[r][c] = cell.getStringCellValue();
					System.out.println(cell.getStringCellValue());
					break;
					
					case NUMERIC:
					obj[r][c]=(int)cell.getNumericCellValue();
					System.out.println(cell.getNumericCellValue());
					break;
				}
				
				c++;
			}
			
			r++;
			c=0;
		}
		
		return obj;
		
		
	}
	
	@DataProvider(name="supplier")
	public Object[][] dataSupplier() {
		
		String excelFilePath = System.getProperty("user.dir")+"\\src\\test\\resources\\tutorialsninjadata.xlsx";
		File excelFile = new File(excelFilePath);
		FileInputStream fis = null;
		
		try {
			fis = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rowsCount = sheet.getPhysicalNumberOfRows(); 
		int colsCount = sheet.getRow(0).getLastCellNum(); 
		
		Object[][] data = new Object[rowsCount-1][colsCount];
		
		for(int r=0;r<rowsCount-1;r++) {  
			
			XSSFRow row = sheet.getRow(r+1);
			
			for(int c=0;c<colsCount;c++) {
				
				XSSFCell cell = row.getCell(c);
				
				CellType cellType = cell.getCellType();
				
				switch(cellType) {
				
				case STRING:
					data[r][c] = cell.getStringCellValue();
					break;
				
				case NUMERIC:
					data[r][c] = (int)cell.getNumericCellValue();
					break;
					
				case BOOLEAN:
					data[r][c] = cell.getBooleanCellValue();
					break;
				
				}
				
			}
			
		}
		
		return data;
		
	}


}
