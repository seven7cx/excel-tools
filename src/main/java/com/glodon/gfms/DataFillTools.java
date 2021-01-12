package com.glodon.gfms;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

        Map<String, Integer> nameMap = new HashMap<>() {
            {
                put("收入", 5);
                put("内部费用-资源采购", 22);
                put("内部收入-资源收入", 8);
                put("内部费用-佣金", 21);
                put("内部收入-佣金", 7);
                put("内部费用-特许经营权", 24);
            }
        };
        nameMap.forEach((sheetName, rowIndex) -> {
            Cell cell = profitSheet.getRow(rowIndex).getCell(11);
            if (workbook.getSheet(sheetName) != null) {
                int lastRowNum = workbook.getSheet(sheetName).getLastRowNum() + 1;
                cell.setCellFormula(String.format("D%d-'%s'!E%d", rowIndex + 1, sheetName, lastRowNum));
            } else {
                cell.setCellValue("");
                cell.setCellFormula(null);
            }
        });

        //产值公式
        Cell cell = summarySheet.getRow(5).getCell(3);
        if (workbook.getSheet("产值") != null) {
            int outputLastRowNum = workbook.getSheet("产值").getLastRowNum() + 1;
            cell.setCellFormula(String.format("产值!E%d/10^4", outputLastRowNum));
        } else {
            cell.setCellFormula(null);
            cell.setCellValue(0);
        }

        List<String> costNameList = Arrays.asList("人力成本", "专项费用", "基本费用");
        for (int i = 0; i < costNameList.size(); i++) {
            Cell costCell = profitSheet.getRow(15 + i).getCell(11);
            costCell.setCellFormula(String.format("D%d-%s!F3", 16 + i, costNameList.get(i)));
        }

        summarySheet.setForceFormulaRecalculation(true);
        profitSheet.setForceFormulaRecalculation(true);
    }
}
