package TESTRUNNER;

import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


import ExtentReport.ExtentTestManager;
import Pages.APIUrl;
import Pages.DemoAPI;
import Utilities.ReadExcel;
import Utilities.Retry;
import Utilities.Utility;
import Utilities.XlsxExcel;
import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;



public class APIRunner {
	String scriptDirectory,APILogss;
	@BeforeSuite
	public void preconditions() {
		scriptDirectory = System.getProperty("user.dir"); 
		APILogss = Utility.createAPIDir(scriptDirectory, "DEMOAPI");    
	    
	  }
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
		
		
         if(XlsxExcel.getData(XlsxExcel.sheetName, "Method").contains("Get")) {
        
		Response res =APIRequest_Response.GETRequest(APIUrl.Geturl);
		String response= res.getBody().prettyPrint();
		
		//Utility.APILog(XlsxExcel.currentTestcase, "Get_Request", APILogss, request);
		Utility.APILog(XlsxExcel.currentTestcase, "Get_Response", APILogss, response);
         }
         if(XlsxExcel.getData(XlsxExcel.sheetName, "Method").contains("Post")) {
          DemoAPI policy=new DemoAPI();
          String request=policy.setPlaceHolder(XlsxExcel.sheetName);
		  Response res =APIRequest_Response.POSTRequest(request, APIUrl.Posturl);
		  String response= res.getBody().prettyPrint();
		 Utility.APILog(XlsxExcel.currentTestcase, "Post_Request", APILogss, request);
		 Utility.APILog(XlsxExcel.currentTestcase, "Post_Response", APILogss, response);
               }
		}
		
		
	
	
	@AfterMethod
	public void Report(ITestResult result) throws Exception {
		XlsxExcel.Result(result);
		
	}
}
