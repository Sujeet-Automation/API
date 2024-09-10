package Utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;

import org.apache.poi.ss.usermodel.Font;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ExcelDataAccess {
	private final String filePath;

	private String datasheetName;

	public String getDatasheetName() {
		return datasheetName;
	}

	public void setDatasheetName(String datasheetName) {
		this.datasheetName = datasheetName;
	}

	public void setValue(int rowNum, String columnHeader, String value) throws Exception {
		setValue(rowNum, columnHeader, value, null);
	}

	public void setValueReport(int rowNum, String columnHeader, String value) throws Exception {
		setValueReport(rowNum, columnHeader, value, null);
	}

	public ExcelDataAccess(String filePath) {
		this.filePath = filePath;

	}

	private void checkPreRequisites() throws Exception {
		if (datasheetName == null) {
			throw new Exception("ExcelDataAccess.datasheetName is not set!");
		}
	}

	private XSSFWorkbook openFileForReading() throws Exception {

		String absoluteFilePath = filePath;

		FileInputStream fileInputStream;
		try {
			fileInputStream = new FileInputStream(absoluteFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new Exception("The specified file \"" + absoluteFilePath + "\" does not exist!");
		}

		XSSFWorkbook workbook;
		try {
			workbook = new XSSFWorkbook(fileInputStream);
			// fileInputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
			throw new Exception("Error while opening the specified Excel workbook \"" + absoluteFilePath + "\"");
		}

		return workbook;
	}

	private void writeIntoFile(XSSFWorkbook workbook) throws Exception {
		String absoluteFilePath = filePath;

		FileOutputStream fileOutputStream;
		try {
			fileOutputStream = new FileOutputStream(absoluteFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new Exception("The specified file \"" + absoluteFilePath + "\" does not exist!");
		}

		try {
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
			throw new Exception("Error while writing into the specified Excel workbook \"" + absoluteFilePath + "\"");
		}
		
//		fileOutputStream.close();
	}

	private XSSFSheet getWorkSheet(XSSFWorkbook workbook) throws Exception {
		XSSFSheet worksheet = workbook.getSheet(datasheetName);
		if (worksheet == null) {
			throw new Exception(
					"The specified sheet \"" + datasheetName + "\"" + "does not exist within the workbook \"");
		}

		return worksheet;
	}

	public int getRowNum(String key, int columnNum, int startRowNum) throws Exception {
		checkPreRequisites();

		XSSFWorkbook workbook = openFileForReading();
		XSSFSheet worksheet = getWorkSheet(workbook);
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

		String currentValue;
		for (int currentRowNum = startRowNum; currentRowNum <= worksheet.getLastRowNum(); currentRowNum++) {

			XSSFRow row = worksheet.getRow(currentRowNum);
			XSSFCell cell = row.getCell(columnNum);
			currentValue = getCellValueAsString(cell, formulaEvaluator);

			if (currentValue.equals(key)) {
				return currentRowNum;
			}
		}

		return -1;
	}

	private String getCellValueAsString(XSSFCell cell, FormulaEvaluator formulaEvaluator) throws Exception {
		if (cell == null || formulaEvaluator.evaluate(cell).getCellType() == CellType.BLANK) {
			return "";
		} else {
			if (formulaEvaluator.evaluate(cell).getCellType() == CellType.ERROR) {
				throw new Exception("Error in formula within this cell! " + "Error code: " + cell.getErrorCellValue());
			}

			DataFormatter dataFormatter = new DataFormatter();
			return dataFormatter.formatCellValue(formulaEvaluator.evaluateInCell(cell));
		}
	}

	public String getValue(int rowNum, String columnHeader) throws Exception {
		checkPreRequisites();

		XSSFWorkbook workbook = openFileForReading();
		XSSFSheet worksheet = getWorkSheet(workbook);
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

		XSSFRow row = worksheet.getRow(0); // 0 because header is always in the
											// first row
		int columnNum = -1;
		String currentValue;
		for (int currentColumnNum = 0; currentColumnNum < row.getLastCellNum(); currentColumnNum++) {

			XSSFCell cell = row.getCell(currentColumnNum);
			currentValue = getCellValueAsString(cell, formulaEvaluator);

			if (currentValue.equals(columnHeader)) {
				columnNum = currentColumnNum;
				break;
			}
		}

		if (columnNum == -1) {
			throw new Exception("The specified column header \"" + columnHeader + "\"" + "is not found in the sheet \""
					+ datasheetName + "\"!");
		} else {
			row = worksheet.getRow(rowNum);
			XSSFCell cell = row.getCell(columnNum);
			return getCellValueAsString(cell, formulaEvaluator);
		}
	}

	public void setValue(int rowNum, String columnHeader, String value, ExcelCellFormatting cellFormatting)
			throws Exception {
		checkPreRequisites();

		XSSFWorkbook workbook = openFileForReading();
		XSSFSheet worksheet = getWorkSheet(workbook);
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

		XSSFRow row = worksheet.getRow(0); // 0 because header is always in the
											// first row
		int columnNum = -1;
		String currentValue;
		for (int currentColumnNum = 0; currentColumnNum < row.getLastCellNum(); currentColumnNum++) {

			XSSFCell cell = row.getCell(currentColumnNum);
			currentValue = getCellValueAsString(cell, formulaEvaluator);

			if (currentValue.equals(columnHeader)) {
				columnNum = currentColumnNum;
				break;
			}
		}

		if (columnNum == -1) {
			throw new Exception("The specified column header \"" + columnHeader + "\"" + "is not found in the sheet \""
					+ datasheetName + "\"!");
		} else {
			row = worksheet.getRow(rowNum);
			XSSFCell cell = row.createCell(columnNum);
			cell.setCellType(CellType.STRING);
			cell.setCellValue(value);

			if (cellFormatting != null) {
				XSSFCellStyle cellStyle = applyCellStyle(workbook, cellFormatting);
				cell.setCellStyle(cellStyle);
			}

			writeIntoFile(workbook);
		}
	}

	public void setValueReport(int rowNum, String columnHeader, String value, ExcelCellFormatting cellFormatting)
			throws Exception {
		checkPreRequisites();

		XSSFWorkbook workbook = openFileForReading();
		XSSFSheet worksheet = getWorkSheet(workbook);
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

		XSSFRow row = worksheet.getRow(0); // 0 because header is always in the
											// first row
		int columnNum = -1;
		String currentValue;
		for (int currentColumnNum = 0; currentColumnNum < row.getLastCellNum(); currentColumnNum++) {

			XSSFCell cell = row.getCell(currentColumnNum);
			currentValue = getCellValueAsString(cell, formulaEvaluator);

			if (currentValue.equals(columnHeader)) {
				columnNum = currentColumnNum;
				break;
			}
		}

		if (columnNum == -1) {
			throw new Exception("The specified column header \"" + columnHeader + "\"" + "is not found in the sheet \""
					+ datasheetName + "\"!");
		} else {

			row = worksheet.getRow(rowNum);
			XSSFCell cell = row.createCell(columnNum);
			cell.setCellType(CellType.STRING);
			// System.out.println("Value is === "+value);
			
			
			
			if (value == "PASS") {
				CellStyle style = workbook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setAlignment(HorizontalAlignment.CENTER);
				Font font = workbook.createFont();
				font.setColor(IndexedColors.WHITE1.getIndex());
				font.setBold(true);
				style.setFont(font);
				cell.setCellStyle(style);
				
			}
			if (value == "FAIL") {
				CellStyle style = workbook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED1.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setAlignment(HorizontalAlignment.CENTER);
				Font font = workbook.createFont();
				font.setColor(IndexedColors.WHITE1.getIndex());
				font.setBold(true);
				style.setFont(font);
				cell.setCellStyle(style);
			}
			cell.setCellValue(value);

			if (cellFormatting != null) {
				XSSFCellStyle cellStyle = applyCellStyle(workbook, cellFormatting);
				cell.setCellStyle(cellStyle);
			}
			writeIntoFile(workbook);
		}

	}

	private XSSFCellStyle applyCellStyle(XSSFWorkbook workbook, ExcelCellFormatting cellFormatting) {
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		if (cellFormatting.centred) {
			cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
		}
		cellStyle.setFillForegroundColor(cellFormatting.getBackColorIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFFont font = workbook.createFont();
		font.setFontName(cellFormatting.getFontName());
		font.setFontHeightInPoints(cellFormatting.getFontSize());
		if (cellFormatting.bold) {
			font.setBold(true);
		}
		font.setColor(cellFormatting.getForeColorIndex());
		cellStyle.setFont(font);

		return cellStyle;
	}
}