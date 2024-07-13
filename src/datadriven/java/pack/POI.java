package datadriven.java.pack;

//POI DataDriven Framework----> XSSF--->new Excel , HSSF---> Old Excel
//JXL---> old only

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class POI {
	public static void main(String[] args) throws IOException, InterruptedException {
		System.setProperty("webdriver.chrome.driver","E:\\Eclipse_Work_Space\\DataDriven-Java-POI\\Drivers\\chromedriver-126.exe");
		WebDriver d = new ChromeDriver();
		d.manage().window().maximize();
		d.get("http://brm.tremplintech.in/web_pages/login.aspx");	
//Import Excel sheet.	
		File src = new File("E:\\Eclipse_Work_Space\\DataDriven-Java-POI\\Supporting Documents\\DD-Java-POI-TestData.xlsx");
//Load the file
		FileInputStream FI = new FileInputStream(src);
//Load Workbook		 
		XSSFWorkbook Workbook = new XSSFWorkbook(FI);
//Load the sheet in which data is stored.
		XSSFSheet sheet = Workbook.getSheetAt(0);
		Cell cell;
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			cell = sheet.getRow(i).getCell(0);
			// cell.setCellType(Cell.getCellType());
			d.findElement(By.id("txt_unam")).sendKeys(cell.getStringCellValue());
			cell = sheet.getRow(i).getCell(1);
			// cell.setCellType(Cell.getCellType());
			d.findElement(By.id("txt_pass")).sendKeys(cell.getStringCellValue());
			d.findElement(By.id("Button3")).click();// login
			Thread.sleep(5000);
			d.findElement(By.xpath("//*[@id=\"LinkButton1\"]")).click(); // Logout
		}
	}
}