package com.glodon.gfms;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.glodon.CommonUtils;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * @author zhangjf-a
 * 01.05.2021
 */
@Slf4j
public class SheetCopyTools {

    @Setter
    private static String filePath = "";
    private static final Map<Integer, XSSFCellStyle> commonStyleMap = new HashMap<>();

    //收入、产值表
    @SneakyThrows
    public static void copyIOSheet(String origFileName, XSSFWorkbook targetWorkbook, String department) {
        XSSFSheet targetSheet1 = targetWorkbook.createSheet("收入");
        XSSFSheet targetSheet2 = targetWorkbook.createSheet("产值");
        //标题行
        Arrays.asList("收入", "产值").forEach(title -> createTitleRow(targetWorkbook.getSheet(title), Arrays.asList("月份", "销售部门", "产品名称", "区域名称", "金额", "归属部门")));

        XSSFCellStyle cellStyle = targetWorkbook.createCellStyle();
        cellStyle.cloneStyleFrom(commonStyleMap.get(targetWorkbook.hashCode()));

        ExcelReader excelReader = EasyExcelFactory.read(filePath + origFileName, null, new AnalysisEventListener<LinkedHashMap<Integer, String>>() {
            @Override
            public void invoke(LinkedHashMap<Integer, String> data, AnalysisContext context) {
                Row row;
                if (data.get(24).contains(department)) {
                    if (data.get(6).contains("收入")) {
                        row = targetSheet1.createRow(targetSheet1.getLastRowNum() + 1);
                    } else {
                        row = targetSheet2.createRow(targetSheet2.getLastRowNum() + 1);
                    }
                } else {
                    return;
                }

                int columnIndex = 0;
                row.createCell(columnIndex++).setCellValue(data.get(19));
                row.createCell(columnIndex++).setCellValue(data.get(3));
                row.createCell(columnIndex++).setCellValue(data.get(8));
                row.createCell(columnIndex++).setCellValue(data.get(12));

                Cell moneyCell = row.createCell(columnIndex++);
                cellStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("#,##0.00"));
                moneyCell.setCellStyle(cellStyle);
                moneyCell.setCellValue(Double.parseDouble(data.get(17)));
                row.createCell(columnIndex).setCellValue(CommonUtils.parseNameFromDisplayName(data.get(24)));
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {

            }
        }).headRowNumber(1).excelType(ExcelTypeEnum.XLSX).build();

        ReadSheet sheet = EasyExcel.readSheet(0).build();
        excelReader.read(sheet);
        //合计行
        Arrays.asList("收入", "产值").forEach(title -> createSumRow(targetWorkbook.getSheet(title)));
        excelReader.finish();
    }

    private static void createTitleRow(XSSFSheet sheet, List<String> titleList) {
        byte[] color = new byte[]{(byte) 220, (byte) 230, (byte) 241};
        XSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.cloneStyleFrom(commonStyleMap.get(sheet.getWorkbook().hashCode()));
        cellStyle.setFillForegroundColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        Row row = sheet.createRow(0);
        for (int i = 0; i < titleList.size(); i++) {
            row.createCell(i).setCellValue(titleList.get(i));
            row.getCell(i).setCellStyle(cellStyle);
        }
    }

    //内部收入
    public static void copyInnerSheet(String origFileName, XSSFWorkbook targetWorkbook, String department) throws IOException {
        XSSFWorkbook origWorkbook = new XSSFWorkbook(new FileInputStream(filePath + origFileName));
        Row row;
        XSSFSheet origSheet = origWorkbook.getSheetAt(0);
        XSSFSheet targetSheet2 = targetWorkbook.createSheet("内部收入-佣金");
        XSSFSheet targetSheet3 = targetWorkbook.createSheet("内部收入-特许经营权");
        XSSFSheet targetSheet5 = targetWorkbook.createSheet("内部费用-佣金");
        XSSFSheet targetSheet6 = targetWorkbook.createSheet("内部费用-特许经营权");

        //标题行
        Arrays.asList("内部收入-佣金", "内部收入-特许经营权", "内部费用-佣金", "内部费用-特许经营权")
            .forEach(title -> createTitleRow(targetWorkbook.getSheet(title), Arrays.asList("月份", "支出部门", "类型", "资源名称", "金额", "比例", "收入部门")));

        XSSFSheet targetSheet;
        XSSFCellStyle moneyCellStyle = targetWorkbook.createCellStyle();
        moneyCellStyle.cloneStyleFrom(commonStyleMap.get(targetWorkbook.hashCode()));
        XSSFCellStyle rateCellStyle = targetWorkbook.createCellStyle();
        rateCellStyle.cloneStyleFrom(commonStyleMap.get(targetWorkbook.hashCode()));
        for (int rowIndex = 3; rowIndex < origSheet.getPhysicalNumberOfRows(); rowIndex++) {
            //create row in this new sheet
            Row origRow = origSheet.getRow(rowIndex);
            String type = origRow.getCell(3).getStringCellValue();
            if (origRow.getCell(4).getStringCellValue().contains(department)) {
                if ("渠道佣金类".equals(type)) {
                    targetSheet = targetSheet2;
                } else if ("特许经营权使用费或综合服务费".equals(type)) {
                    targetSheet = targetSheet3;
                } else {
                    continue;
                }
            } else if (origRow.getCell(9).getStringCellValue().contains(department)) {
                if ("渠道佣金类".equals(type)) {
                    targetSheet = targetSheet5;
                } else if ("特许经营权使用费或综合服务费".equals(type)) {
                    targetSheet = targetSheet6;
                } else {
                    continue;
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
            moneyCellStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("#,##0.00"));
            moneyCellStyle.setAlignment(HorizontalAlignment.RIGHT);
            moneyCell.setCellStyle(moneyCellStyle);
            moneyCell.setCellValue(origRow.getCell(8).getNumericCellValue());

            Cell rateCell = row.createCell(columnIndex++);
            rateCellStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("0.00%"));
            rateCellStyle.setAlignment(HorizontalAlignment.RIGHT);
            rateCell.setCellStyle(rateCellStyle);
            rateCell.setCellValue(origRow.getCell(7).getNumericCellValue());
            row.createCell(columnIndex).setCellValue(origRow.getCell(4).getStringCellValue());
        }

        //合计行
        Arrays.asList("内部收入-佣金", "内部收入-特许经营权", "内部费用-佣金", "内部费用-特许经营权")
            .forEach(title -> createSumRow(targetWorkbook.getSheet(title)));
        origWorkbook.close();
    }

    private static void createSumRow(Sheet sheet) {
        XSSFCellStyle commonStyle = commonStyleMap.get(sheet.getWorkbook().hashCode());
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            for (int j = 0; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (CellType.STRING.equals(cell.getCellType())) {
                    cell.setCellStyle(commonStyle);
                }
            }
        }

        int lastRowIndex = sheet.getLastRowNum() + 1;
        if (lastRowIndex == 1) {
            sheet.getWorkbook().removeSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
        }
        sheet.addMergedRegion(new CellRangeAddress(lastRowIndex, lastRowIndex, 0, 3));
        sheet.createRow(lastRowIndex).createCell(0).setCellValue("合计");
        sheet.getRow(lastRowIndex).getCell(0).setCellStyle(commonStyle);
        sheet.getRow(lastRowIndex).createCell(1).setCellStyle(commonStyle);
        sheet.getRow(lastRowIndex).createCell(2).setCellStyle(commonStyle);
        sheet.getRow(lastRowIndex).createCell(3).setCellStyle(commonStyle);

        Cell cell = sheet.getRow(lastRowIndex).createCell(4);
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.cloneStyleFrom(commonStyle);
        cellStyle.setDataFormat(sheet.getWorkbook().createDataFormat().getFormat("#,##0.00"));
        cell.setCellStyle(cellStyle);
        String columnLetter = CellReference.convertNumToColString(cell.getColumnIndex());
        cell.setCellFormula(String.format("SUM(%s:%s)", columnLetter + 2, columnLetter + lastRowIndex));
        sheet.setForceFormulaRecalculation(true);

        for (int i = 0; i < sheet.getRow(0).getPhysicalNumberOfCells(); i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    //资源交易
    public static void copyResourceSheet(String origFileName, XSSFWorkbook targetWorkbook, String department) throws IOException {
        XSSFWorkbook origWorkbook = new XSSFWorkbook(new FileInputStream(filePath + origFileName));
        Row row;
        XSSFSheet origSheet = origWorkbook.getSheetAt(0);
        XSSFSheet targetSheet1 = targetWorkbook.createSheet("内部收入-资源收入");
        XSSFSheet targetSheet2 = targetWorkbook.createSheet("内部费用-资源采购");
        //标题行
        Arrays.asList("内部收入-资源收入", "内部费用-资源采购").forEach(title -> createTitleRow(targetWorkbook.getSheet(title), Arrays.asList("月份", "支出部门", "类型", "资源名称", "金额", "收入部门")));

        XSSFSheet targetSheet;
        XSSFCellStyle numberStyle = targetWorkbook.createCellStyle();
        numberStyle.cloneStyleFrom(commonStyleMap.get(targetWorkbook.hashCode()));
        for (int rowIndex = 1; rowIndex < origSheet.getPhysicalNumberOfRows(); rowIndex++) {
            //create row in this new sheet
            Row origRow = origSheet.getRow(rowIndex);
            if (origRow.getCell(15).getStringCellValue().contains(department)) {
                targetSheet = targetSheet1;
            } else if (origRow.getCell(17).getStringCellValue().contains(department)) {
                targetSheet = targetSheet2;
            } else {
                continue;
            }

            row = targetSheet.createRow(targetSheet.getLastRowNum() + 1);
            int columnIndex = 0;
            int month = (int) (origRow.getCell(23).getNumericCellValue());
            row.createCell(columnIndex++).setCellValue(String.format("%d-%02d", month / 100, month % 100));
            row.createCell(columnIndex++).setCellValue(CommonUtils.parseNameFromDisplayName(origRow.getCell(17).getStringCellValue()));
            row.createCell(columnIndex++).setCellValue(origRow.getCell(7).getStringCellValue());
            row.createCell(columnIndex++).setCellValue(origRow.getCell(8).getStringCellValue());

            Cell moneyCell = row.createCell(columnIndex++);
            numberStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat("#,##0.00"));
            moneyCell.setCellStyle(numberStyle);
            moneyCell.setCellValue(origRow.getCell(1).getNumericCellValue());
            row.createCell(columnIndex).setCellValue(CommonUtils.parseNameFromDisplayName(origRow.getCell(15).getStringCellValue()));
        }

        //合计行
        Arrays.asList("内部收入-资源收入", "内部费用-资源采购").forEach(title -> createSumRow(targetWorkbook.getSheet(title)));
        origWorkbook.close();
    }

    public static void initCommonStyle(XSSFWorkbook workbook) {
        XSSFCellStyle commonStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeight((short) 200);
        font.setFontName("微软雅黑");
        commonStyle.setFont(font);
        commonStyle.setBorderTop(BorderStyle.THIN);
        commonStyle.setBorderBottom(BorderStyle.THIN);
        commonStyle.setBorderLeft(BorderStyle.THIN);
        commonStyle.setBorderRight(BorderStyle.THIN);
        commonStyleMap.put(workbook.hashCode(), commonStyle);
    }

    //人专基费用
    public static void copyCostSheet(String origFileName, XSSFWorkbook targetWorkbook, String department) throws IOException {
        XSSFWorkbook origWorkbook = new XSSFWorkbook(new FileInputStream(filePath + origFileName));
        XSSFRow row;
        XSSFSheet origSheet = origWorkbook.getSheetAt(0);
        XSSFSheet targetSheet = targetWorkbook.createSheet(formatSheetName(origFileName));
        //标题行
        copyTitleRow(origSheet, targetSheet);
        XSSFCellStyle stringStyle = targetWorkbook.createCellStyle();
        XSSFCellStyle numberStyle = targetWorkbook.createCellStyle();
        XSSFCellStyle rateStyle = targetWorkbook.createCellStyle();
        for (int rowIndex = 5; rowIndex < origSheet.getPhysicalNumberOfRows(); rowIndex++) {
            if (!department.equals(origSheet.getRow(rowIndex).getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue())) {
                //且排除部门名称不匹配的行
                continue;
            }
            //create row in this new sheet
            row = targetSheet.createRow(targetSheet.getLastRowNum() + 1);
            doCopySheetData(origSheet, row, rowIndex, stringStyle, numberStyle, rateStyle);
        }
        createSumRowAndStyle(targetSheet);
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

    private static void copyTitleRow(XSSFSheet origSheet, XSSFSheet targetSheet) {
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
        targetSheet.addMergedRegion(new CellRangeAddress(0, 1, 3, 3));

        XSSFCellStyle titleCellStyle = targetSheet.getWorkbook().createCellStyle();
        doCopySheetTitle(origSheet, targetSheet.createRow(0), 2, titleCellStyle);
        doCopySheetTitle(origSheet, targetSheet.createRow(1), 3, titleCellStyle);
        XSSFCellStyle summaryCellStyle = targetSheet.getWorkbook().createCellStyle();
        XSSFCellStyle summaryNumberCellStyle = targetSheet.getWorkbook().createCellStyle();
        XSSFCellStyle summaryRateCellStyle = targetSheet.getWorkbook().createCellStyle();
        doCopySheetData(origSheet, targetSheet.createRow(2), 4, summaryCellStyle, summaryNumberCellStyle, summaryRateCellStyle);
    }

    private static void doCopySheetTitle(XSSFSheet origSheet, XSSFRow row, int rowIndex, XSSFCellStyle titleCellStyle) {
        for (int colIndex = 0; colIndex < origSheet.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
            XSSFCell origCell = origSheet.getRow(rowIndex).getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            XSSFCell targetCell = row.createCell(colIndex);
            XSSFCellStyle origStyle = origCell.getCellStyle();
            titleCellStyle.cloneStyleFrom(origStyle);
            targetCell.setCellStyle(titleCellStyle);
            switch (origCell.getCellType()) {
                case STRING -> targetCell.setCellValue(origCell.getRichStringCellValue().getString());
                case NUMERIC -> targetCell.setCellValue(origCell.getNumericCellValue());
                case FORMULA -> targetCell.setCellValue(origCell.getCellFormula());
                case BLANK -> targetCell.setBlank();
                default -> log.warn("No matched cell Type of {}", origCell.getCellType().toString());
            }
        }
    }

    private static void doCopySheetData(XSSFSheet origSheet, XSSFRow row, int rowIndex, XSSFCellStyle stringStyle, XSSFCellStyle numberStyle, XSSFCellStyle rateStyle) {
        for (int colIndex = 0; colIndex < origSheet.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
            XSSFCell origCell = origSheet.getRow(rowIndex).getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            XSSFCell targetCell = row.createCell(colIndex);
            XSSFCellStyle origStyle = origCell.getCellStyle();
            switch (origCell.getCellType()) {
                case STRING -> {
                    targetCell.setCellValue(origCell.getRichStringCellValue().getString());
                    stringStyle.cloneStyleFrom(origStyle);
                    targetCell.setCellStyle(stringStyle);
                }
                case NUMERIC -> {
                    targetCell.setCellValue(origCell.getNumericCellValue());
                    if (targetCell.getColumnIndex() % 3 == 0) {
                        rateStyle.cloneStyleFrom(origStyle);
                        targetCell.setCellStyle(rateStyle);
                    } else {
                        numberStyle.cloneStyleFrom(origStyle);
                        targetCell.setCellStyle(numberStyle);
                    }
                }
                case FORMULA -> targetCell.setCellValue(origCell.getCellFormula());
                case BLANK -> {
                    targetCell.setBlank();
                    targetCell.setCellStyle(numberStyle);
                }
                default -> log.warn("No matched cell Type of {}", origCell.getCellType().toString());
            }
        }
    }

    private static void createSumRowAndStyle(XSSFSheet sheet) {
        final int lastRowIndex = sheet.getLastRowNum();
        if (lastRowIndex == 3) {
            sheet.getWorkbook().removeSheetAt(sheet.getWorkbook().getSheetIndex(sheet));
            return;
        }

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

        Map<Integer, Integer> maxWidth = new HashMap<>(sheet.getRow(0).getPhysicalNumberOfCells());
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < sheet.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
                if (CellType.STRING.equals(sheet.getRow(rowIndex).getCell(colIndex).getCellType())) {
                    maxWidth.put(colIndex, Math.max(maxWidth.getOrDefault(colIndex, 0), sheet.getRow(rowIndex).getCell(colIndex).getStringCellValue().length() * 2 * 256));
                    sheet.setColumnWidth(colIndex, maxWidth.get(colIndex));
                } else if (rowIndex > 0 && colIndex % 3 == 0) {
                    sheet.setColumnWidth(colIndex, 3500);
                } else if (CellType.NUMERIC.equals(sheet.getRow(rowIndex).getCell(colIndex).getCellType())) {
                    sheet.setColumnWidth(colIndex, 3500);
                }
            }
        }
    }
}