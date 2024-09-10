package Utilities;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Random;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;



public class Utility {
	static int i = 0 ;
	static String lastName="";
	static String firstName="";
	static String middleName="";
	static String mobileNum="";
	
	static String sDate = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
	static Calendar cal = Calendar.getInstance();
	static SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
	static String sTime = sdf.format(cal.getTime());

	
	public static String generateFirstName(String userExisting) {

		if(userExisting.equals("")) {
			String AlphabeticString="abcdefghijklmnopqrstuvwxyz";
			int nLen=6;
			StringBuilder sb= new StringBuilder(nLen);
			for (int i = 0; i < nLen; i++) {
				int index=(int) (AlphabeticString.length()*Math.random());
				sb.append(AlphabeticString.charAt(index));
			}
			firstName =sb.toString();
		}
		return firstName;
	}
	
	public static String generateMiddleName(String userExisting) {

		if(userExisting.equals("")) {
			String AlphabeticString="abcdefghijklmnopqrstuvwxyz";
			int nLen=6;
			StringBuilder sb= new StringBuilder(nLen);
			for (int i = 0; i < nLen; i++) {
				int index=(int) (AlphabeticString.length()*Math.random());
				sb.append(AlphabeticString.charAt(index));
			}
			middleName =sb.toString();
		}
		return middleName;
	}

	public static String generateLastName(String userExisting) {

		if(userExisting.equals("")) {
			String AlphabeticString="abcdefghijklmnopqrstuvwxyz";
			int nLen=6;
			StringBuilder sb=new StringBuilder(nLen);
			for (int i = 0; i < nLen; i++) {
				int index=(int) (AlphabeticString.length() * Math.random());
				sb.append(AlphabeticString.charAt(index));
			}
			lastName=sb.toString();
		}
		return lastName;
	}

	public static String generateEmail() {
		return firstName+"."+lastName+"@gmail.com";
	}

	public static String generateMobileNum(String userExisting) {

		if(userExisting.equals("")) {
			String NumericString="1234567890";
			int nLen=10;
			StringBuilder sb= new StringBuilder(nLen);
			sb.append(5);
			for (int i = 1; i < nLen; i++) {
				int index=(int) (NumericString.length()*Math.random());
				sb.append(NumericString.charAt(index));
			}
			mobileNum =sb.toString();
		}
		return mobileNum;
	}

	public static String createAPIDir(String APILogss,String ProductName) {
		String sDate = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		System.out.println(sdf.format(cal.getTime()));
		String sTime = sdf.format(cal.getTime());
		APILogss = APILogss + File.separator + "APILogs"+File.separator +ProductName+ File.separator + sDate.replace("-", "_") + "_" + sTime.replace(":", "_");
		try {
			if (!new File(APILogss).exists()) {
				new File(APILogss).mkdirs();
			}
		} catch (Exception e) {
			//LOGGER.error("Screenshot Path Exception:"+e.getMessage());
		}
		return APILogss;
	}

	public static String createScreenshotDir(String screenshotpath,String ProductName) {
		
		screenshotpath = screenshotpath + File.separator + "Reports_And_Screenshots"+File.separator +ProductName+ File.separator + sDate.replace("-", "_") + "_" + sTime.replace(":", "_");
		try {
			if (!new File(screenshotpath).exists()) {
				new File(screenshotpath).mkdirs();
			}
		} catch (Exception e) {
			//LOGGER.error("Screenshot Path Exception:"+e.getMessage());
		}
		return screenshotpath;
	}
	public static void Screenshot(WebDriver driver, String TestCaseId, String Name,String APILogss) {

		APILogss = APILogss + File.separator + TestCaseId;
		try {
			if (!new File(APILogss).exists()) {
				new File(APILogss).mkdirs();
			}
		} catch (Exception e) {
			
		}

		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		String dDate = new SimpleDateFormat("HH:mm:ss.SSS").format(new Date());
		dDate = dDate.replace(":", "").replace(" ", "").replace(".", "");

		try {
			FileUtils.copyFile(src, new File(APILogss + "/" + dDate + "_" + Name + ".png"));
		} catch (IOException e) {
			
			e.printStackTrace();
		}
	}
	
	public static void APILog(String TestCaseId, String Name,String ScreenShotPath,String APIBody) throws IOException {

		ScreenShotPath = ScreenShotPath + File.separator + TestCaseId;
		try {
			if (!new File(ScreenShotPath).exists()) {
				new File(ScreenShotPath).mkdirs();
			}
		} catch (Exception e) {
			//LOGGER.error("Screenshot Path Exception:"+e.getMessage());
		}

		//File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		String dDate = new SimpleDateFormat("HH:mm:ss.SSS").format(new Date());
		dDate = dDate.replace(":", "").replace(" ", "").replace(".", "");
		
		FileWriter fw=new FileWriter(ScreenShotPath + "/"  + Name + ".json");    
        fw.write(APIBody);    
        fw.close();

		
	}
	public static void CarinaReport(String Path) throws Exception {
		Path = Path + File.separator + "CarinaReport";
		Path ="./Reports_And_Screenshots/Endorsement/"+ sDate.replace("-", "_") + "_" + sTime.replace(":", "_")+"/CarinaReport" ;
		try {
			if (!new File(Path).exists()) {
				new File(Path).mkdirs();
			}
		} catch (Exception e) {
			//LOGGER.error("Screenshot Path Exception:"+e.getMessage());
		}
		
		
	}

	
	
	public static String generateRandomChars(String candidateChars, int length) {
		StringBuilder sb = new StringBuilder();
		Random random = new Random();
		for (int i = 0; i < length; i++) {
			sb.append(candidateChars.charAt(random.nextInt(candidateChars.length())));
		}
		return sb.toString();
	}
}

