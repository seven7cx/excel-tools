package com.glodon.gfms;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author zhangjf-a
 * 01.06.2021
 */
public class DataFillTools {

    public static void fillSummarySheet(String origFileName, XSSFWorkbook workbook, String department) throws IOException {
        XSSFSheet summarySheet = workbook.getSheetAt(0);
        summarySheet.getRow(1).getCell(0).setCellValue("部门：" + department);

        XSSFSheet profitSheet = workbook.getSheetAt(1);
        profitSheet.getRow(1).getCell(1).setCellValue("部门：" + department);
        XSSFSheet origSheet = new XSSFWorkbook(new FileInputStream(origFileName)).getSheetAt(0);
        int origRowIndex = 0;
        for (int i = 0; i < origSheet.getPhysicalNumberOfRows(); i++) {
            if (department.equals(origSheet.getRow(i).getCell(0).getStringCellValue())) {
                origRowIndex = i + 1;
                break;
            }
        }
        for (int i = 5; i < profitSheet.getPhysicalNumberOfRows(); i++, origRowIndex++) {
            Cell cell = profitSheet.getRow(i).getCell(2);
            if (cell.getCellStyle().getFillForegroundColor() == 0) {
                cell.setCellValue(origSheet.getRow(origRowIndex).getCell(3).getNumericCellValue());
                profitSheet.getRow(i).getCell(3).setCellValue(origSheet.getRow(origRowIndex).getCell(4).getNumericCellValue());
            }
        }
    }
}
