package attendance;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

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
public class Punch_In_Out {
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
	public void validationApp() {	
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
		HSSFSheet s1 = wb1.getSheet("Sheet4");
		WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));	
		Actions act=new Actions(driver);
		act.moveToElement(time).perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement attendance = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));	
		Actions act1=new Actions(driver); 
		act1.moveToElement(attendance).click().perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement punchin = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
		act1.moveToElement(punchin).click().perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement attendancedate = driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue()));
		attendancedate.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement month = driver.findElement(By.xpath(s1.getRow(5).getCell(1).getStringCellValue()));	
		Select sel=new Select(month);	
		sel.selectByValue("5");	
		WebElement year = driver.findElement(By.xpath(s1.getRow(6).getCell(1).getStringCellValue()));	
		Select sel1=new Select(year);	
		sel1.selectByValue("2023");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		String searchdate = "3";
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
		WebElement in = driver.findElement(By.xpath(s1.getRow(8).getCell(1).getStringCellValue()));	
		in.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement attendancedate1 = driver.findElement(By.xpath(s1.getRow(9).getCell(1).getStringCellValue()));
		attendancedate1.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement month11 = driver.findElement(By.xpath(s1.getRow(10).getCell(1).getStringCellValue()));	
		Select sel12=new Select(month11);	
		sel12.selectByValue("5");	
		WebElement year11 = driver.findElement(By.xpath(s1.getRow(11).getCell(1).getStringCellValue()));	
		Select sel11=new Select(year11);	
		sel11.selectByValue("2023");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		String searchdate11 = "4";
		List<WebElement> alldates11 = driver.findElements(By.xpath(s1.getRow(12).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates11)
		{
			String dt=elements.getText();
			
		   if (dt.equals(searchdate11))
		    {
		       elements.click(); 
		       break;
		    }
		 		}
		WebElement out = driver.findElement(By.xpath(s1.getRow(13).getCell(1).getStringCellValue()));	
		out.click();
		}
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
		catch (Exception e) {
		File f11 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);		
		try {
		Files.copy(f11, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\employeerecords.png"));	
		} catch (IOException e1) {
			e1.printStackTrace();
}
}
}
}	
		
		
		
		
		
