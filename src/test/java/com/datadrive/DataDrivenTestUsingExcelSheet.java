package com.datadrive;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenTestUsingExcelSheet {

	WebDriver driver;

	@BeforeClass
	public void setup() {

		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofMillis(5000));

	}

	@Test(dataProvider = "readData")
	public void loginTest(String loginID, String pass, String stat) {
		driver.get("https://admin-demo.nopcommerce.com/");
		WebElement id = driver.findElement(By.xpath("//*[@id=\"Email\"]"));
		id.clear();
		id.sendKeys(loginID);

		WebElement password = driver.findElement(By.xpath("//input[@id='Password']"));
		password.clear();
		password.sendKeys(pass);
		driver.findElement(By.xpath("//button[contains(text(),'Log in')]")).click();

		String actualTitle = driver.getTitle();
		String expectedTitele = "Dashboard / nopCommerce administration";

		if (stat.equals("Valid")) {
			if (actualTitle.equals(expectedTitele)) {

				driver.findElement(By.xpath("//a[contains(text(), 'Logout')]")).click();
				Assert.assertTrue(true);

			} else {
				Assert.assertTrue(false);
			}
		} else if (stat.equals("Invalid")) {
			if (actualTitle.equals(expectedTitele)) {
				driver.findElement(By.xpath("//a[contains(text(), 'Logout')]")).click();
				Assert.assertTrue(false);

			} else {

				Assert.assertTrue(true);

			}
		}
	}

	@DataProvider(name = "readData")
	public String[][] readData() throws IOException {

//				String loginData[][] = { 
//				{ "admin@yourstore.com", "admin", "Valid" },
//				{ "admin@yourstore.com", "adm", "Invalid" }, 
//				{ "adm@yorstore.com", "admin", "Invalid" },
//				{ "adm@yourstore.com", "adm", "Invalid" } 
//				};

		// ExcelUtil util=new
		// ExcelUtil("C:\\Users\\emadu\\eclipse-workspace\\Selenium_Excel_Sheet\\src\\test\\resources\\TestData_Excel\\TEstData.xlsx");
		// ExcelUtil util=new
		// ExcelUtil(".\\src\\test\\resources\\TestData_Excel\\TEstData.xlsx");
		ExcelUtil util = new ExcelUtil(
		System.getProperty("user.dir") + "\\src\\test\\resources\\TestData_Excel\\TEstData.xlsx");
		int totalRows = util.getRowCount("Sheet1");
		int totalCells = util.getCellCount("Sheet1", 1);

		String loginData[][] = new String[totalRows][totalCells];

		for (int i = 1; i <= totalRows; i++) {
			for (int j = 0; j < totalCells; j++) {

				loginData[i - 1][j] = util.getCellData("Sheet1", i, j);
			}
		}
		return loginData;
	}

	@AfterClass
	public void teardown() {

		driver.close();

	}

}
