package attendance;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import jxl.Sheet;
import jxl.Workbook;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.google.common.io.Files;

import bussiop.Pom;

public class My_Records {
	public static WebDriver driver;
	public Workbook wb; 
	public Sheet sh;
	@BeforeTest 	
	public void launchApp()  {					
		System.setProperty("webdriver.chrome.driver", "D:\\ECLIPS\\ORANGE_HRM(TIME)\\ChromeDriver\\chromedriver.exe"); 													
		driver = new ChromeDriver();													
		driver.get("https://adminjon-osondemand.orangehrm.com/symfony/web/index.php/auth/login");													
		driver.manage().window().maximize();													
	}
	@AfterTest		
	public void closeApp() { 		
	driver.close();		
	}		
	@Test						
	public void validationApp(){		
		try {
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		Pom p =PageFactory.initElements(driver,Pom.class);
		p.uid.sendKeys(s.getCell(1, 5).getContents());
		p.password.sendKeys(s.getCell(1, 6).getContents());
		p.sub.click();
		FileInputStream f1 = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\Attendance.xls");	
		try (HSSFWorkbook wb1 = new HSSFWorkbook(f1)) {
			HSSFSheet s1 = wb1.getSheet("Sheet3");
			WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));	
			Actions act=new Actions(driver);
			Thread.sleep(3000);	
			act.moveToElement(time).perform();
			WebElement attendance = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));	
			Actions act1=new Actions(driver); 
			act1.moveToElement(attendance).click().perform();
			WebElement myrecords = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
			act1.moveToElement(myrecords).click().perform();
			WebElement attendancedate = driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue()));
			attendancedate.click();
			WebElement month = driver.findElement(By.xpath(s1.getRow(5).getCell(1).getStringCellValue()));	
			Select sel=new Select(month);	
			sel.selectByValue("4");	
			WebElement year = driver.findElement(By.xpath(s1.getRow(6).getCell(1).getStringCellValue()));	
			Select sel1=new Select(year);	
			sel1.selectByValue("2023");
			Thread.sleep(3000);	
			String searchdate = "25";
			List<WebElement> alldates = driver.findElements(By.xpath(s1.getRow(7).getCell(1).getStringCellValue()));
			for (WebElement elements : alldates)
			{
				String dt=elements.getText();
				
			    if (dt.equals(searchdate))
			    {
			        elements.click(); 
			        break;
			    }
			 		}
		}
		Thread.sleep(10000);
		}
		catch (Exception e) {
		File f11 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	
		try {
		Files.copy(f11, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\myrecords.png"));	
		} catch (IOException e1) {
			e1.printStackTrace();
		}	
	}
}
}
