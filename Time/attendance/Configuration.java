package attendance;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;				
import org.apache.poi.hssf.usermodel.HSSFWorkbook;				
import java.io.File;		
import com.google.common.io.Files;
import bussiop.Pom;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.PageFactory;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
public class Configuration extends My_Records {
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
	public void validationApp()  {		
		try {
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		Pom p =PageFactory.initElements(driver,Pom.class);
		p.getUid().sendKeys(s.getCell(1, 5).getContents());
		p.getPassword().sendKeys(s.getCell(1, 6).getContents());
		p.getSub().click();
		FileInputStream f1 = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\Attendance.xls");
		try (HSSFWorkbook wb1 = new HSSFWorkbook(f1)) {
			HSSFSheet s1 = wb1.getSheet("Sheet1");				
			WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));
			Actions act=new Actions(driver);	
			act.moveToElement(time).perform();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			WebElement attendance = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));
			Actions act1=new Actions(driver); 	
			act1.moveToElement(attendance).click().perform();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);	
			WebElement config = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
			act1.moveToElement(config).click().perform();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			WebElement save = driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue()));
			save.click();
		}
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		}
			catch (Exception e) {
		File f12 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);											
		try {
			Files.copy(f12, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\configuration.png"));
		} catch (IOException e1) {
			e1.printStackTrace();
		}											
		}
}
}
