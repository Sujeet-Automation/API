package Utilities;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {
    public static String[] ExcelData(String path) {
        List<String> stringList = new ArrayList<>();
        try {
            Workbook workbook;
            FileInputStream fs = new FileInputStream(path);

            workbook = new XSSFWorkbook(fs);
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet != null) {
                int rowCount = sheet.getLastRowNum();
                for (int i = 1; i <= rowCount; i++) {
                    Row curRow = sheet.getRow(i);

                    if (curRow == null) {
                        rowCount = i - 1;
                        break;
                    }

                    String firstCellValue = curRow.getCell(0) != null ? curRow.getCell(0).toString() : null;
                    if (firstCellValue == null || firstCellValue.isEmpty()) {
                        continue;
                    }

                    String executeval = curRow.getCell(0).getStringCellValue().trim();
                    if (executeval.equals("Y")) {
                        String testCaseID = curRow.getCell(1).getStringCellValue().trim();
                        stringList.add(testCaseID);
                    }
                }
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return stringList.toArray(new String[0]);
    }
}

