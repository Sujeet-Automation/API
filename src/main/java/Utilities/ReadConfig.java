package Utilities;

import java.io.File;
import java.io.FileInputStream;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class ReadConfig {
	
	Properties pro;

	public ReadConfig ()  {

	File src=new File("./Configuration/Configuration.properties");
	try {
	FileInputStream fis = new FileInputStream(src);
	
	pro=new Properties();
	pro.load(fis);
	} catch (Exception e) {
		
		e.printStackTrace();
	}
	
	}
	public String getUrL() {
		String url=pro.getProperty("baseURL");
		return url;
	}
}
