package Pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;

import Utilities.Utility;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;
	public static String scriptDirectory, screenshotpath;

	@BeforeClass
	public void setup() throws Exception {
		scriptDirectory = System.getProperty("user.dir");
		screenshotpath = Utility.createScreenshotDir(scriptDirectory, "API");
		Utility.CarinaReport(screenshotpath);
	}

}
