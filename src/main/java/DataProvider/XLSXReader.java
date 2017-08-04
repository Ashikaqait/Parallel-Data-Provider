package DataProvider;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class XLSXReader {

	

	@DataProvider(name = "empLogin",parallel=true)
	public Object[][] loginData() throws BiffException, IOException {
		Object[][] arrayObject = getExcelData("D:\\test\\DataProvider\\src\\main\\resources\\dp.xls", "Sheet1");

		return arrayObject;
	}

	@Test(dataProvider = "empLogin")

	public void test_mail(String recipient, String subject, String text) throws InterruptedException {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\ashikasrivastava\\Downloads\\chromedriver_win32 (2)\\chromedriver.exe");

		WebDriver driver = new ChromeDriver();

		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		driver.get("https://accounts.google.com/ServiceLogin?");

		driver.findElement(By.id("identifierId")).sendKeys("ms2894907@gmail.com");
		driver.findElement(By.id("identifierNext")).click();

		Thread.sleep(3000);

		driver.findElement(By.xpath(".//*[@id='password']/div[1]/div/div[1]/input")).sendKeys("Abc@1234");
		driver.findElement(By.id("passwordNext")).click();
		driver.findElement(By.cssSelector("a.WaidBe")).click();
	

		driver.findElement(By.xpath(".//*[@id=':4s']/div/div")).click();
		Thread.sleep(2000);

		 driver.findElement(By.xpath("//textarea[@name='to']")).sendKeys(recipient);
		driver.findElement(By.xpath("//input[@placeholder='Subject']")).sendKeys(subject);
		driver.findElement(By.xpath("//div[@aria-label='Message Body']")).sendKeys(text);
		driver.findElement(By.xpath("//div[text()='Send']")).click();

	}

	public String[][] getExcelData(String fileName, String sheetName) throws BiffException, IOException {
		String[][] arrayExcelData = null;

		FileInputStream fs = new FileInputStream(fileName);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet(sheetName);

		int Column = sh.getColumns();
		int Rows = sh.getRows();

		arrayExcelData = new String[Rows][Column];
	
		for (int i = 0; i < Rows; i++) {

			for (int j = 0; j < Column; j++) {
				arrayExcelData[i][j] = sh.getCell(j, i).getContents();
			}

		}

		return arrayExcelData;
	}
}
