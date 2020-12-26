package TestExcelsheet;

import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import ExcelUtility.XLUtils;
import io.github.bonigarcia.wdm.WebDriverManager;

public class WorldTableDataTest {
	
   static WebDriver driver;
	
	public static void main(String[] args) throws Exception {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
		
		
		//create excel sheet
		String path = ".\\datafiles\\population.xlsx";
		
		//write header in excel sheet
		
		XLUtils.setCellData(path, "Sheet1", 0, 0, "Country");
		XLUtils.setCellData(path, "Sheet1", 0, 1, "Population");
		XLUtils.setCellData(path, "Sheet1", 0, 2, "% of world");
		XLUtils.setCellData(path, "Sheet1", 0, 3, "Date");
		XLUtils.setCellData(path, "Sheet1", 0, 4, "Source");
		
		//capture table rows
		
		WebElement table = driver.findElement(By.xpath("//table[@class=\"wikitable sortable plainrowheaders jquery-tablesorter\"]/tbody"));
		int rows = table.findElements(By.xpath("tr")).size();  //rows present in web Table
		
		for(int r=1; r<=rows; r++)
		{
			//read data from web table
			String country           = table.findElement(By.xpath("tr["+r+"]/td[1]")).getText();
			String population        = table.findElement(By.xpath("tr["+r+"]/td[2]")).getText();
			String percentageofworld = table.findElement(By.xpath("tr["+r+"]/td[3]")).getText();
			String date              = table.findElement(By.xpath("tr["+r+"]/td[4]")).getText();
			String source            = table.findElement(By.xpath("tr["+r+"]/td[5]")).getText();
			
			System.out.println(country +"---"+ population +"---"+ percentageofworld +"---"+ date +"---"+ source);
			
			//writing the data in excel sheet
			
			XLUtils.setCellData(path, "Sheet1", r, 0, country);
			XLUtils.setCellData(path, "Sheet1", r, 1, population);
			XLUtils.setCellData(path, "Sheet1", r, 2, percentageofworld);
			XLUtils.setCellData(path, "Sheet1", r, 3, date);
			XLUtils.setCellData(path, "Sheet1", r, 4, source);
		
		}
		
		
		System.out.println("Data Added in Excel Sheet Succesfully!!!!!!!");
		driver.close();
	}

}
