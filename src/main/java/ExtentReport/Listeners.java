package ExtentReport;

import static ExtentReport.ExtentManager.takeScreenshot;

import java.io.File;

import org.openqa.selenium.WebDriver;
import org.testng.ITestListener;
import org.testng.ITestResult;

import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;


public class Listeners implements ITestListener{ 
	//public static String currentTC;
    @Override
    public synchronized void onTestStart(ITestResult result) {
        System.out.println("--------- Executing :- " + getSimpleMethodName(result) + " ---------");
        //ExtentTestManager.createTest(result.getName(),result.getMethod().getDescription());
        ExtentTestManager.setCategoryName(getSimpleClassName(result));
    }

    @Override
    public synchronized void onTestSuccess(ITestResult result) {
        ExtentTestManager.getTest().assignCategory(getSimpleClassName(result));
        addExtentLabelToTest(result);
        ExtentTestManager.endTest();
    }

    @Override
    public synchronized void onTestFailure(ITestResult result) {
        String concatenate=".";
        ExtentTestManager.getTest().assignCategory(getSimpleClassName(result));
        ExtentTestManager.getTest().log(Status.FAIL, "<b>"+result.getName() + "</b> Test has failed</br><font color= red>" +result.getThrowable()+"</font>");
        String testMethodName=result.getMethod().getMethodName();
        System.out.println(testMethodName);

        WebDriver driver=null;

        try {
            driver=(WebDriver)result.getTestClass().getRealClass().getDeclaredField("driver").get(result.getInstance());
        } catch (Exception e) {
            e.printStackTrace();
        }

        try {
            ExtentTestManager.getTest().fail(MediaEntityBuilder.createScreenCaptureFromPath(concatenate+takeScreenshot(testMethodName,driver)).build());
        } catch (Exception e) {
            e.printStackTrace();
        }
        addExtentLabelToTest(result);
        ExtentTestManager.endTest();
    }

    @Override
    public synchronized void onTestSkipped(ITestResult result) {
        ExtentTestManager.getTest().log(Status.SKIP, result.getName() + " Test has Skipped<br/><font color= orange>" + result.getThrowable()+"</font>");
    }


    private synchronized String getSimpleClassName(ITestResult result) {
        return result.getMethod().getRealClass().getSimpleName();
    }

    private synchronized String getSimpleMethodName(ITestResult result) {
        return result.getName();
    }

    private synchronized void addExtentLabelToTest(ITestResult result) {
        if (result.getStatus() == ITestResult.SUCCESS)
            ExtentTestManager.getTest().pass(MarkupHelper.createLabel("Test Passed", ExtentColor.GREEN));
        else if (result.getStatus() == ITestResult.FAILURE) {
            ExtentTestManager.getTest().fail(MarkupHelper.createLabel("Test Failed", ExtentColor.RED));
        } else
            ExtentTestManager.getTest().skip(MarkupHelper.createLabel("Test Skipped", ExtentColor.ORANGE));
    }
    
    
    public static void logPass(String message) {
    	ExtentTestManager.getTest().log(Status.PASS,"<font color= green>"+message+"</font>");
    }
    
    public static void logFail(String message) {
    	ExtentTestManager.getTest().log(Status.FAIL,"<font color= red>"+message+"</font>");
    }
}

