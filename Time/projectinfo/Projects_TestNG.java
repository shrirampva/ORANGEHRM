package projectinfo;
import java.io.File;
import java.io.FileInputStream;
import jxl.Sheet;
import jxl.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
//import org.openqa.selenium.Keys;
//import org.openqa.selenium.OutputType;
//import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
//import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.google.common.io.Files;
public class Projects_TestNG {
	public static WebDriver driver;
	public Workbook wb; 
	public Sheet sh;
	@BeforeTest 	
	public void LaunchApp() throws Exception, Throwable {					
		System.setProperty("webdriver.chrome.driver", "D:\\ECLIPS\\ORANGE_HRM(TIME)\\ChromeDriver\\chromedriver.exe"); 													
		driver = new ChromeDriver();													
		driver.get("https://adminjon-osondemand.orangehrm.com/symfony/web/index.php/auth/login");													
		driver.manage().window().maximize();													
	}
	@AfterTest		
	public void CloseApp() { 		
	driver.close();		
	}		
	@Test						
	public void ValidationApp() throws Throwable, Exception {						
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		driver.findElement(By.name(s.getCell(1, 2).getContents())).sendKeys(s.getCell(1, 5).getContents());											
		driver.findElement(By.name(s.getCell(1, 3).getContents())).sendKeys(s.getCell(1, 6).getContents());											
		driver.findElement(By.name(s.getCell(1, 4).getContents())).click();
		WebElement time = driver.findElement(By.xpath("//*[@id=\"menu_time_viewTimeModule\"]/b"));	
		Actions act=new Actions(driver);	
		act.moveToElement(time).perform();
		Thread.sleep(3000);	
		WebElement infoproj = driver.findElement(By.xpath("//*[@id=\"menu_admin_ProjectInfo\"]"));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(infoproj).click().perform();
		WebElement proj = driver.findElement(By.xpath("//*[@id=\"menu_admin_viewProjects\"]"));	
		act1.moveToElement(proj).click().perform();
		WebElement name = driver.findElement(By.xpath("//*[@id=\"searchProject_customer\"]"));	
		name.clear();	
		Thread.sleep(3000);	
		name.sendKeys("MARY");
		WebElement pro = driver.findElement(By.xpath("//*[@id=\"searchProject_project\"]"));	
		pro.clear();	
		pro.sendKeys("S");
		WebElement proadmin = driver.findElement(By.xpath("//*[@id=\"searchProject_projectAdmin\"]"));	
		proadmin.clear();	
		Thread.sleep(3000);	
		proadmin.sendKeys("Jon Philip");
		WebElement add = driver.findElement(By.xpath("//*[@id=\"btnSearch\"]"));	
		add.click();	
		Thread.sleep(10000);	
		File f1 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);											
		Files.copy(f1, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\project.png"));											

		
		
		
	}
}
