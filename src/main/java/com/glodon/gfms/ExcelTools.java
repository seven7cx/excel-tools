package com.glodon.gfms;

import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author zhangjf-a
 * 01.05.2021
 */
@Slf4j
public class ExcelTools {

    @Setter
    private static String filePath = "";

    public static void copyInnerSheet(String origFileName, XSSFWorkbook targetWorkbook, String department) throws IOException {
        XSSFWorkbook origWorkbook = new XSSFWorkbook(new FileInputStream(filePath + origFileName));
        Row row;
        XSSFSheet origSheet = origWorkbook.getSheetAt(0);
        XSSFSheet targetSheet2 = targetWorkbook.createSheet("内部收入-佣金");
        XSSFSheet targetSheet3 = targetWorkbook.createSheet("内部收入-特许经营权");
        XSSFSheet targetSheet5 = targetWorkbook.createSheet("内部费用-佣金");
        XSSFSheet targetSheet6 = targetWorkbook.createSheet("内部费用-特许经营权");

        Map<String, Double> sumMap = new HashMap<>() {
            {
                put("内部收入-佣金", 0.0);
                put("内部收入-特许经营权", 0.0);
                put("内部费用-佣金", 0.0);
                put("内部费用-特许经营权", 0.0);
            }
        };

        //标题行
        sumMap.forEach((k, v) -> createTitleRow(targetWorkbook.getSheet(k), targetWorkbook.createCellStyle(), targetWorkbook.createFont()));

        XSSFSheet targetSheet;
        XSSFCellStyle cellStyle = targetWorkbook.createCellStyle();
        for (int rowIndex = 3; rowIndex < origSheet.getPhysicalNumberOfRows(); rowIndex++) {
            //create row in this new sheet
            Row origRow = origSheet.getRow(rowIndex);
            String type = origRow.getCell(3).getStringCellValue();
            if (origRow.getCell(4).getStringCellValue().contains(department)) {
                if ("PC与PC之间内部交易(012)".equals(type)) {
                    continue;
                } else if ("渠道佣金类".equals(type)) {
                    targetSheet = targetSheet2;
                } else {
                    targetSheet = targetSheet3;
                }
            } else if (origRow.getCell(9).getStringCellValue().contains(department)) {
                if ("PC与PC之间内部交易(012)".equals(type)) {
                    continue;
                } else if ("渠道佣金类".equals(type)) {
                    targetSheet = targetSheet5;
                } else {
                    targetSheet = targetSheet6;
                }
            } else {
                continue;
            }

            row = targetSheet.createRow(targetSheet.getLastRowNum() + 1);
            int columnIndex = 0;
            row.createCell(columnIndex++).setCellValue(origRow.getCell(1).getStringCellValue());
            row.createCell(columnIndex++).setCellValue(origRow.getCell(9).getStringCellValue());
            row.createCell(columnIndex++).setCellValue(type);
            row.createCell(columnIndex++).setCellValue(origRow.getCell(6).getStringCellValue());

            Cell moneyCell = row.createCell(columnIndex++);
            cellStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("#,##0.00"));
            moneyCell.setCellStyle(cellStyle);
            moneyCell.setCellValue(origRow.getCell(8).getNumericCellValue());

            Cell rateCell = row.createCell(columnIndex++);
            cellStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("0%"));
            rateCell.setCellStyle(cellStyle);
            rateCell.setCellValue(origRow.getCell(7).getNumericCellValue());
            row.createCell(columnIndex).setCellValue(origRow.getCell(4).getStringCellValue());

            sumMap.put(targetSheet.getSheetName(), sumMap.get(targetSheet.getSheetName()) + origRow.getCell(8).getNumericCellValue());
        }

        //合计行
        sumMap.forEach((k, v) -> createSumRow(targetWorkbook.getSheet(k), v, cellStyle));
        origWorkbook.close();
    }

    private static void createTitleRow(Sheet sheet, XSSFCellStyle cellStyle, Font font) {
        byte[] color = new byte[]{(byte) 220, (byte) 230, (byte) 241};
        cellStyle.setFillForegroundColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        font.setFontName("黑体");
        cellStyle.setFont(font);
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("月份");
        row.createCell(1).setCellValue("支出部门");
        row.createCell(2).setCellValue("类型");
        row.createCell(3).setCellValue("资源名称");
        row.createCell(4).setCellValue("金额");
        row.createCell(5).setCellValue("比例");
        row.createCell(6).setCellValue("收入部门");
        for (int i = 0; i < 7; i++) {
            sheet.setColumnWidth(i, 20 * 256);
            sheet.getRow(0).getCell(i).setCellStyle(cellStyle);
        }
    }

    private static void createSumRow(Sheet sheet, double sum, CellStyle cellStyle) {
        int lastRowIndex = sheet.getLastRowNum() + 1;
        sheet.addMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 0, 3));
        sheet.createRow(lastRowIndex).createCell(0).setCellValue("合计");
        Cell cell = sheet.getRow(lastRowIndex).createCell(4);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(sum);
    }

    public static void copySheet(String origFileName, XSSFWorkbook targetWorkbook, String department) throws IOException {
        XSSFWorkbook origWorkbook = new XSSFWorkbook(new FileInputStream(filePath + origFileName));
        XSSFCellStyle newStyle = targetWorkbook.createCellStyle();
        Row row;
        XSSFSheet origSheet = origWorkbook.getSheetAt(0);
        XSSFSheet targetSheet = targetWorkbook.createSheet(formatSheetName(origFileName));
        //标题行
        copyTitleRow(origSheet, targetSheet, newStyle);
        for (int rowIndex = 5; rowIndex < origSheet.getPhysicalNumberOfRows(); rowIndex++) {
            if (!department.equals(origSheet.getRow(rowIndex).getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue())) {
                //且排除部门名称不匹配的行
                continue;
            }
            //create row in this new sheet
            row = targetSheet.createRow(targetSheet.getLastRowNum() + 1);
            doCopySheetData(origSheet, row, rowIndex, newStyle);
        }
        createSumRowAndStyle(targetWorkbook, targetSheet);
        origWorkbook.close();
    }

    private static String formatSheetName(String origFileName) {
        return switch (origFileName) {
            case "0401.人力成本预实.xlsx" -> "人力成本";
            case "0402.专项费用预实.xlsx" -> "专项费用";
            case "0403.基本费用预实.xlsx" -> "基本费用";
            default -> "其它";
        };
    }

    private static void copyTitleRow(XSSFSheet origSheet, XSSFSheet targetSheet, XSSFCellStyle newStyle) {
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 3, 3));

        doCopySheetData(origSheet, targetSheet.createRow(0), 2, newStyle);
        doCopySheetData(origSheet, targetSheet.createRow(1), 3, newStyle);
        doCopySheetData(origSheet, targetSheet.createRow(2), 4, newStyle);
    }

    private static void doCopySheetData(XSSFSheet origSheet, Row row, int rowIndex, XSSFCellStyle newStyle) {
        for (int colIndex = 0; colIndex < origSheet.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
            XSSFCell origCell = origSheet.getRow(rowIndex).getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell targetCell = row.createCell(colIndex);
            CellStyle origStyle = origCell.getCellStyle();
            newStyle.cloneStyleFrom(origStyle);
            targetCell.setCellStyle(newStyle);
            switch (origCell.getCellType()) {
                case STRING -> targetCell.setCellValue(origCell.getRichStringCellValue().getString());
                case NUMERIC -> targetCell.setCellValue(origCell.getNumericCellValue());
                case FORMULA -> targetCell.setCellValue(origCell.getCellFormula());
                case BLANK -> targetCell.setBlank();
                default -> log.warn("No matched cell Type of {}", origCell.getCellType().toString());
            }
        }
    }

    private static void createSumRowAndStyle(XSSFWorkbook workbook, XSSFSheet sheet) {
        final int lastRowIndex = sheet.getLastRowNum();
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 3));
        Row sumRow = sheet.getRow(2);
        sumRow.getCell(0).setCellValue("合计");
        for (int i = 4; i < 43; i++) {
            Cell cell = sumRow.getCell(i);
            if (i % 3 == 0) {
                cell.setCellFormula(String.format("%s/%s", CellReference.convertNumToColString(i - 1) + 3, CellReference.convertNumToColString(i - 2) + 3));
            } else {
                String columnLetter = CellReference.convertNumToColString(cell.getColumnIndex());
                cell.setCellFormula(String.format("SUM(%s:%s)", columnLetter + 4, columnLetter + (lastRowIndex + 1)));
            }
        }

        sheet.setForceFormulaRecalculation(true);

        XSSFCellStyle origStyle = sheet.getRow(0).getCell(0).getCellStyle();
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.cloneStyleFrom(origStyle);
        byte[] color = new byte[]{(byte) 220, (byte) 230, (byte) 241};
        cellStyle.setFillForegroundColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle rateStyle = workbook.createCellStyle();
        rateStyle.cloneStyleFrom(origStyle);
        rateStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < sheet.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
                if (CellType.STRING.equals(sheet.getRow(rowIndex).getCell(colIndex).getCellType())) {
                    sheet.getRow(rowIndex).getCell(colIndex).setCellStyle(cellStyle);
                } else if (rowIndex > 1 && colIndex % 3 == 0) {
                    sheet.getRow(rowIndex).getCell(colIndex).setCellStyle(rateStyle);
                }
            }
        }
    }
}
