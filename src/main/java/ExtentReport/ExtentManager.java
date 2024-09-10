package ExtentReport;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import com.aventstack.extentreports.ExtentReports;

import com.aventstack.extentreports.reporter.configuration.Theme;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExtentManager {
   private static ExtentReports extent;
//public static  String OUTPUT_FOLDER_SCREENSHOTS;
//public static  String REPORT_FILE_PATH;
 public static final String OUTPUT_FOLDER_SCREENSHOTS ="./FailedScreenshots/";
 public static final String REPORT_FILE_PATH ="./ExtentReport/";

   public static String dateGenerator()
   {
       DateFormat dateFormat = new SimpleDateFormat("dd_MMM_yyyy_HH_mm_ss");
       Date date = new Date();
       return dateFormat.format(date);
   }

   public static void createInstance()
   {
       String path=System.getProperty("user.dir")+REPORT_FILE_PATH;
       ExtentSparkReporter reporter = new ExtentSparkReporter(path+"Test-Automaton-Report"+"_"+dateGenerator()+".html");
       reporter.config().setReportName("Automation Results");
       reporter.config().setDocumentTitle("Test Results");
       //reporter.config().setTheme(Theme.DARK);
       extent=new ExtentReports();
       extent.attachReporter(reporter);
       extent.setSystemInfo("Tester","Sujeet");
       extent.setSystemInfo("OS", System.getProperty("os.name"));
       extent.setSystemInfo("Java",  System.getProperty("java.specification.version"));
   }

   public static ExtentReports getInstance() {
       if (extent == null)
           createInstance();
       return extent;
   }

   public static synchronized String takeScreenshot(String methodName,WebDriver driver) throws IOException {

       String filePathExtent = OUTPUT_FOLDER_SCREENSHOTS +  methodName + "_" + dateGenerator() + ".png";
       String filePath = ExtentManager.REPORT_FILE_PATH + filePathExtent;
       try {
           TakesScreenshot ts = (TakesScreenshot) driver;
           File source = ts.getScreenshotAs(OutputType.FILE);
           FileUtils.copyFile(source, new File(filePath));
       }
       catch (IOException e){
           e.getStackTrace();
           Reporter.log("Failed To Take screenshot " + e, true);
       }
       return filePath;
   }
}