package stock;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.Test;

import Utilities.XLUtility;
import io.github.bonigarcia.wdm.WebDriverManager;

public class StockAnalysis {
	WebDriver driver;
	String sectorName;
	int r, r1, size;

	@Test
	public void sectorName() throws IOException, Throwable {

		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
		driver.get("https://www.screener.in/explore/");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		String path = ".\\resources\\stocks.xlsx";
		File xlfile = new File(path);
		if (xlfile.exists()) {
			xlfile.delete();
		}
		// int sectorsSize = driver.findElements(By.xpath("//aside/div/div/a")).size();
		for (size = 1; size <= 85; size++) {
			WebElement sector = driver.findElement(By.xpath("//aside/div/div/a[" + size + "]"));
			System.out.println(size + ". " + sector.getText());

			if (sector.getText().length() > 28) {
				sectorName = sector.getText().substring(0, 28).replaceAll("/", " ");
			} else {
				sectorName = sector.getText().replaceAll("/", " ");
				;
			}
			sector.click();

			WebElement table = driver.findElement(By.xpath("//table[1]/tbody[1]"));
			XLUtility xlutil = new XLUtility(path);

			// Write headers in excel sheet

			xlutil.setCellData(sectorName, 0, 0, "S.No.");
			xlutil.setCellData(sectorName, 0, 1, "Name");
			xlutil.setCellData(sectorName, 0, 2, "CMP Rs.");
			xlutil.setCellData(sectorName, 0, 3, "P/E");
			xlutil.setCellData(sectorName, 0, 4, "Mar Cap Rs.Cr.");
			xlutil.setCellData(sectorName, 0, 5, "Div Yld %");
			xlutil.setCellData(sectorName, 0, 6, "NP Qtr Rs.Cr.");
			xlutil.setCellData(sectorName, 0, 7, "Qtr Profit Var %");
			xlutil.setCellData(sectorName, 0, 8, "Sales Qtr Rs.Cr.");
			xlutil.setCellData(sectorName, 0, 9, "Qtr Sales Var %");
			xlutil.setCellData(sectorName, 0, 10, "ROCE %");

			// capture table rows

			int rows = table.findElements(By.xpath("tr")).size(); // rows present in web table
			int count = 0;
			for (r = 2, r1 = 1; r <= rows; r++, r1++) {
				count++;
				if (count != 16) {

					// read data from web table
					table = driver.findElement(By.xpath("//table[1]/tbody[1]"));

					String SNo = table.findElement(By.xpath("tr[" + r + "]/td[1]")).getText();
					String Name = table.findElement(By.xpath("tr[" + r + "]/td[2]")).getText();
					String CMPRs = table.findElement(By.xpath("tr[" + r + "]/td[3]")).getText();
					String PE = table.findElement(By.xpath("tr[" + r + "]/td[4]")).getText();
					String MarCapRsCr = table.findElement(By.xpath("tr[" + r + "]/td[5]")).getText();
					String DivYld = table.findElement(By.xpath("tr[" + r + "]/td[6]")).getText();
					String NPQtrRsCr = table.findElement(By.xpath("tr[" + r + "]/td[7]")).getText();
					String QtrProfitVar = table.findElement(By.xpath("tr[" + r + "]/td[8]")).getText();
					String SalesQtrRsCr = table.findElement(By.xpath("tr[" + r + "]/td[9]")).getText();
					String QtrSalesVar = table.findElement(By.xpath("tr[" + r + "]/td[10]")).getText();
					String ROCE = table.findElement(By.xpath("tr[" + r + "]/td[11]")).getText();
					
					// writing the data in excel sheet
					xlutil.setCellData(sectorName, r1, 0, SNo);
					xlutil.setCellData(sectorName, r1, 1, Name);
					xlutil.setCellData(sectorName, r1, 2, CMPRs);
					xlutil.setCellData(sectorName, r1, 3, PE);
					xlutil.setCellData(sectorName, r1, 4, MarCapRsCr);
					xlutil.setCellData(sectorName, r1, 5, DivYld);
					xlutil.setCellData(sectorName, r1, 6, NPQtrRsCr);
					xlutil.setCellData(sectorName, r1, 7, QtrProfitVar);
					xlutil.setCellData(sectorName, r1, 8, SalesQtrRsCr);
					xlutil.setCellData(sectorName, r1, 9, QtrSalesVar);
					xlutil.setCellData(sectorName, r1, 10, ROCE);

				} else {
					count = 0;
					r1 = r1 - 1;
				}
				try {
					if (r == rows && driver.findElement(By.xpath("//a[normalize-space()='Next']")).isDisplayed()) {
						driver.findElement(By.xpath("//a[normalize-space()='Next']")).click();
						r = 1;
						count = 0;
						rows = driver.findElements(By.xpath("//table[1]/tbody[1]/tr")).size();

					}
				} catch (Exception e) {

				}

			}
			driver.get("https://www.screener.in/explore/");

		}
		System.out.println("All " + size + " Sectors datas are updated successfully");
		//Closing driver
		driver.close();

	}

}
