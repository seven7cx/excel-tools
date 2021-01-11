package com.glodon.gfms;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;

/**
 * ExcelTools Tester.
 *
 * @author zhangjf-a
 * @version 1.0
 * @since 01/05/2021
 */
public class ExcelToolsTest {

    private static XSSFWorkbook targetWorkbook;
    private static final String filePath = "/Users/zhangjingfei/Downloads/temp/excelTools/";
    private static final String department = "计价产品部";
    private static File targetFile;

    @BeforeAll
    public static void beforeAll() throws Exception {
        Assertions.assertTrue(new File(filePath).exists());
        File exampleFile = new File(filePath + "模板.xlsx");
        Assertions.assertTrue(exampleFile.exists());

        targetFile = new File(filePath + "2020年年度结算确认-" + department + ".xlsx");
        if (targetFile.exists()) {
            boolean delete = targetFile.delete();
            Assertions.assertTrue(delete);
            System.out.println("Delete test file " + department);
        }

        Files.copy(exampleFile.toPath(), targetFile.toPath());
        Assertions.assertTrue(targetFile.exists());

        SheetCopyTools.setFilePath(filePath);
        targetWorkbook = new XSSFWorkbook(new FileInputStream(targetFile.getAbsolutePath()));
        SheetCopyTools.initCommonStyle(targetWorkbook);
    }

    @BeforeEach
    public void before() throws Exception {
    }

    @AfterEach
    public void after() throws Exception {
    }

    @AfterAll
    public static void afterAll() throws Exception {
        try (
            FileOutputStream fileOutputStream = new FileOutputStream(targetFile.getAbsolutePath())
        ) {
            targetWorkbook.write(fileOutputStream);
            targetWorkbook.close();
        }
    }

    /**
     * Method: copyIOSheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopyIOSheet() throws IOException {
        SheetCopyTools.copyIOSheet("0505.收入预实表.xlsx", targetWorkbook, department);
    }

    /**
     * Method: copyInnerSheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopyInnerSheet() throws IOException {
        SheetCopyTools.copyInnerSheet("0301.佣金特许权收入.xlsx", targetWorkbook, department);
    }

    /**
     * Method: copyResourceSheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopyResourceSheet() throws IOException {
        SheetCopyTools.copyResourceSheet("bi_resource.xlsx", targetWorkbook, department);
    }

    /**
     * Method: copyCostSheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopyCostSheet() throws IOException {
        SheetCopyTools.copyCostSheet("0402.专项费用预实.xlsx", targetWorkbook, department);
        SheetCopyTools.copyCostSheet("0401.人力成本预实.xlsx", targetWorkbook, department);
        SheetCopyTools.copyCostSheet("0403.基本费用预实.xlsx", targetWorkbook, department);
    }

    /**
     * Method: fillSummarySheet(String origFileName, XSSFWorkbook workbook, String department) throws IOException
     */
    @Test
    public void testFillSummarySheet() throws IOException {
        DataFillTools.fillSummarySheet(filePath + "010102.预实利润表.xlsx", targetWorkbook, department);
    }
}
