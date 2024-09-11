package TESTRUNNER;

import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import ExtentReport.ExtentTestManager;
import Pages.BaseClass;
import Pages.DemoAPI;
import Utilities.ReadExcel;
import Utilities.Retry;
import Utilities.XlsxExcel;
import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.response.Response;



public class APIRunner extends BaseClass{

	@DataProvider

	public String[] Authentication() throws Exception{
	     String[] testObjArray = ReadExcel.ExcelData(scriptDirectory + "/src/test/Receipt_Clear.xlsx");
	     return (testObjArray);
	}
	
	@Test(dataProvider="Authentication",retryAnalyzer = Retry.class)
	public void TestRun(String TestCaseID) throws Exception {
	    
		XlsxExcel.path = scriptDirectory + "/src/test/Receipt_Clear.xlsx";
		XlsxExcel.currentTestcase = TestCaseID;
		XlsxExcel.sheetName = "Sheet1"; 
		

			
		ExtentTestManager.createTest(TestCaseID);
		DemoAPI policy=new DemoAPI();
		String request=policy.setPlaceHolder(XlsxExcel.sheetName);

		Response response =APIRequest_Response.POSTRequest(request, APIUrl.Posturl);
		System.out.println(response.prettyPrint());
		
		Response response1 =APIRequest_Response.GETRequest(APIUrl.Geturl);
		System.out.println(response.prettyPrint());
		System.out.println(response1.prettyPrint());
		}
		
		
	
	
	@AfterMethod
	public void Report(ITestResult result) throws Exception {
		XlsxExcel.Result(result);
		
	}
}
