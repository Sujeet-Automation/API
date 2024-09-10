package Utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.ITestResult;



public class XlsxExcel {
	public static String path;
	public static String currentTestcase;
	public static String sheetName;
	public static String sheetName2;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;
	private static String dataReferenceIdentifier = "#";

	public XlsxExcel(String path) {

		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static String getDataWithSubIteration(String datasheetName, String fieldName, int currentIteration,
			int currentSubIteration) throws Exception {

		ExcelDataAccess testDataAccess = new ExcelDataAccess(path);
		testDataAccess.setDatasheetName(datasheetName);

		int rowNum = testDataAccess.getRowNum(currentTestcase, 0, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The test case \"" + currentTestcase + "\"" + "is not found in the test data sheet \""
					+ datasheetName + "\"!");
		}
		rowNum = testDataAccess.getRowNum(Integer.toString(currentIteration), 1, rowNum);
		if (rowNum == -1) {
			throw new Exception("The iteration number \"" + currentIteration + "\"" + "of the test case \""
					+ currentTestcase + "\"" + "is not found in the test data sheet \"" + datasheetName + "\"!");
		}
	

		String dataValue = testDataAccess.getValue(rowNum, fieldName);

		if (dataValue.startsWith(dataReferenceIdentifier)) {
			dataValue = getCommonData(fieldName, dataValue);
		}

		return dataValue;
	}

	public static String getData(String datasheetName, String fieldName) throws Exception {

		ExcelDataAccess testDataAccess = new ExcelDataAccess(path);
		testDataAccess.setDatasheetName(datasheetName);

		int rowNum = testDataAccess.getRowNum(currentTestcase, 1, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The test case \"" + currentTestcase + "\"" + "is not found in the test data sheet \""
					+ datasheetName + "\"!");
		}

		if (rowNum == -1) {
			throw new Exception("The iteration number \"" + "\"" + "of the test case \"" + currentTestcase + "\""
					+ "is not found in the test data sheet \"" + datasheetName + "\"!");
		}
	

		String dataValue = testDataAccess.getValue(rowNum, fieldName);

		if (dataValue.startsWith(dataReferenceIdentifier)) {
			dataValue = getCommonData(fieldName, dataValue);
		}

		return dataValue;
	}

	private static String getCommonData(String fieldName, String dataValue) throws Exception {
		ExcelDataAccess commonDataAccess = new ExcelDataAccess(path);
		commonDataAccess.setDatasheetName("Common_Testdata");

		String dataReferenceId = dataValue.split(dataReferenceIdentifier)[1];

		int rowNum = commonDataAccess.getRowNum(dataReferenceId, 0, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The common test data row identified by \"" + dataReferenceId + "\""
					+ "is not found in the common test data sheet!");
		}

		return commonDataAccess.getValue(rowNum, fieldName);
	}

	public static void putDatawithSubiteration(String datasheetName, String fieldName, String dataValue,
			int currentIteration, int currentSubIteration) throws Exception {

		ExcelDataAccess testDataAccess = new ExcelDataAccess(path);
		testDataAccess.setDatasheetName(datasheetName);

		int rowNum = testDataAccess.getRowNum(currentTestcase, 0, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The test case \"" + currentTestcase + "\"" + "is not found in the test data sheet \""
					+ datasheetName + "\"!");
		}
		rowNum = testDataAccess.getRowNum(Integer.toString(currentIteration), 1, rowNum);
		if (rowNum == -1) {
			throw new Exception("The iteration number \"" + currentIteration + "\"" + "of the test case \""
					+ currentTestcase + "\"" + "is not found in the test data sheet \"" + datasheetName + "\"!");
		}
	

		synchronized (XlsxExcel.class) {
			testDataAccess.setValue(rowNum, fieldName, dataValue);
		}
	}

	public static void putData(String datasheetName, String fieldName, String dataValue) throws Exception {

		ExcelDataAccess testDataAccess = new ExcelDataAccess(path);
		testDataAccess.setDatasheetName(datasheetName);

		int rowNum = testDataAccess.getRowNum(currentTestcase, 1, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The test case \"" + currentTestcase + "\"" + "is not found in the test data sheet \""
					+ datasheetName + "\"!");
		}

		if (rowNum == -1) {
			throw new Exception("The iteration number \"" + "\"" + "of the test case \"" + currentTestcase + "\""
					+ "is not found in the test data sheet \"" + datasheetName + "\"!");
		}
	

		synchronized (XlsxExcel.class) {
			testDataAccess.setValue(rowNum, fieldName, dataValue);
		}
	}

	public static void SetDataReport(String datasheetName, String fieldName, String Testcase, String dataValue)
			throws Exception {

		ExcelDataAccess testDataAccess = new ExcelDataAccess(path);
		testDataAccess.setDatasheetName(datasheetName);

		int rowNum = testDataAccess.getRowNum(currentTestcase, 1, 1); // Start
																		// at
																		// row
																		// 1,
																		// skipping
																		// the
																		// header
																		// row
		if (rowNum == -1) {
			throw new Exception("The test case \"" + currentTestcase + "\"" + "is not found in the test data sheet \""
					+ datasheetName + "\"!");
		}

		if (rowNum == -1) {
			throw new Exception("The iteration number \"" + "\"" + "of the test case \"" + currentTestcase + "\""
					+ "is not found in the test data sheet \"" + datasheetName + "\"!");
		}
		

		synchronized (XlsxExcel.class) {
			testDataAccess.setValueReport(rowNum, fieldName, dataValue);
		}
	}

	public static String getFileSeparator() {
		return System.getProperty("file.separator");
	}

	public static void Result(ITestResult result) throws Exception {

		long milliseconds = result.getEndMillis() - result.getStartMillis();
		long minutes = (milliseconds / 1000) / 60;
		long seconds = (milliseconds / 1000) % 60;
		String Excutiontime = (minutes + " Min and " + seconds + " sec");
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		Date date = new Date();

		if (ITestResult.SUCCESS == result.getStatus()) {
			XlsxExcel.SetDataReport(sheetName, "Execution_Result", currentTestcase, "PASS");
			XlsxExcel.SetDataReport(sheetName, "Execution_time_Taken", currentTestcase, Excutiontime);
			XlsxExcel.SetDataReport(sheetName, "Execution_Date", currentTestcase, formatter.format(date));
			XlsxExcel.SetDataReport(sheetName, "Failure_Reason", currentTestcase, "");
		} else if (ITestResult.FAILURE == result.getStatus()) {
			Throwable Failure = result.getThrowable();
			String Failure_Reson = String.valueOf(Failure);
			XlsxExcel.SetDataReport(sheetName, "Execution_Result", currentTestcase, "FAIL");
			XlsxExcel.SetDataReport(sheetName, "Execution_time_Taken", currentTestcase, Excutiontime);
			XlsxExcel.SetDataReport(sheetName, "Execution_Date", currentTestcase, formatter.format(date));
			XlsxExcel.SetDataReport(sheetName, "Failure_Reason", currentTestcase, Failure_Reson);
		} else if (ITestResult.SKIP == result.getStatus()) {
			XlsxExcel.SetDataReport(sheetName, "Execution_Result", currentTestcase, "SKIP");
			XlsxExcel.SetDataReport(sheetName, "Execution_time_Taken", currentTestcase, Excutiontime);
			XlsxExcel.SetDataReport(sheetName, "Execution_Date", currentTestcase, formatter.format(date));
		}
	}
	public static void SummaryReport(String Path) throws Exception {
		Path = Path + File.separator ;
		try {
			if (!new File(Path).exists()) {
				new File(Path).mkdirs();
			}
		} catch (Exception e) {
			//LOGGER.error("Screenshot Path Exception:"+e.getMessage());
		}
		File source = new File(path);
		File dest = new File(Path +"ExcelSummary.xlsx");
		try {
		    FileUtils.copyFile(source, dest);
		} catch (IOException e) {
		    e.printStackTrace();
		}
	}
	
	}
