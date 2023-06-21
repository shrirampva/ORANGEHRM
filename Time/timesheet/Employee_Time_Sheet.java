package timesheet;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.google.common.io.Files;
public class Employee_Time_Sheet {
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
		driver.findElement(By.name(s.getCell(1, 2).getContents())).sendKeys(s.getCell(1, 5).getContents());											
		driver.findElement(By.name(s.getCell(1, 3).getContents())).sendKeys(s.getCell(1, 6).getContents());											
		driver.findElement(By.name(s.getCell(1, 4).getContents())).click();
		FileInputStream f1 = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\Timesheet.xls");	
		try (HSSFWorkbook wb1 = new HSSFWorkbook(f1)) {
			HSSFSheet s1 = wb1.getSheet("Sheet2");
		WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));	
		Actions act=new Actions(driver);	
		act.moveToElement(time).perform();
		Thread.sleep(3000);	
		WebElement timeSheets = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(timeSheets).click().perform();
		WebElement employeetimeSheets = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
		act1.moveToElement(employeetimeSheets).click().perform();
		Thread.sleep(3000);	
		driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue())).sendKeys("Stephen Robert");	
		driver.findElement(By.xpath(s1.getRow(5).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(6).getCell(1).getStringCellValue())).sendKeys("Stephen Robert");
		Thread.sleep(3000);
		driver.findElement(By.xpath(s1.getRow(7).getCell(1).getStringCellValue())).click();
		Thread.sleep(3000);	
		String value = "0";
		List<WebElement> alldates = driver.findElements(By.xpath(s1.getRow(8).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates)
		{
			String dt=elements.getText();
		    if (dt.equals(value))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
		WebElement reset = driver.findElement(By.xpath(s1.getRow(9).getCell(1).getStringCellValue()));	
		reset.click();
		String defaultcondition = "edit";
		switch(defaultcondition) {
			default :WebElement edit = driver.findElement(By.xpath(s1.getRow(10).getCell(1).getStringCellValue()));	
			edit.click();	
		    break;
		}	
		Thread.sleep(3000);	
		WebElement row = driver.findElement(By.xpath(s1.getRow(11).getCell(1).getStringCellValue()));	
		row.clear();	
		row.sendKeys("MARY - S");
		row.sendKeys(Keys.ARROW_DOWN);			
		row.sendKeys(Keys.ENTER);			
		Thread.sleep(4000);	
		Select drpActivicty = new Select(driver.findElement(By.name(s1.getRow(12).getCell(1).getStringCellValue())));													
        drpActivicty.selectByVisibleText("ENG");
		Thread.sleep(1000);	
		driver.findElement(By.xpath(s1.getRow(13).getCell(1).getStringCellValue())).sendKeys("8:00");
		driver.findElement(By.xpath(s1.getRow(14).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(15).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(16).getCell(1).getStringCellValue())).sendKeys("8:00");	
		driver.findElement(By.xpath(s1.getRow(17).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(18).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(19).getCell(1).getStringCellValue())).sendKeys("8:00");	
		driver.findElement(By.xpath(s1.getRow(20).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(21).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(22).getCell(1).getStringCellValue())).sendKeys("8:00");	
		driver.findElement(By.xpath(s1.getRow(23).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(24).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(25).getCell(1).getStringCellValue())).sendKeys("8:00");	
		driver.findElement(By.xpath(s1.getRow(26).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(27).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(28).getCell(1).getStringCellValue())).sendKeys("0:00");
		driver.findElement(By.xpath(s1.getRow(29).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(30).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(31).getCell(1).getStringCellValue())).sendKeys("0:00");	
		driver.findElement(By.xpath(s1.getRow(32).getCell(1).getStringCellValue())).clear();
		driver.findElement(By.xpath(s1.getRow(33).getCell(1).getStringCellValue())).sendKeys("9:00");
		driver.findElement(By.xpath(s1.getRow(34).getCell(1).getStringCellValue())).click();	
		Thread.sleep(1000);	
		driver.findElement(By.xpath(s1.getRow(35).getCell(1).getStringCellValue())).click();
		driver.findElement(By.xpath(s1.getRow(36).getCell(1).getStringCellValue())).click();
		driver.findElement(By.xpath(s1.getRow(37).getCell(1).getStringCellValue())).click();
		}
		Thread.sleep(10000);	
		}
		catch (Exception e) {
		File f11 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	
		try {
		Files.copy(f11, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\employeetimesheet.png"));	
		} catch (IOException e1) {
			e1.printStackTrace();
}
	}
	}
}
		