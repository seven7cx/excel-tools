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
    private static final String department = "广材产品部";
    private static File targetFile;

    @BeforeAll
    public static void beforeAll() throws Exception {
        Assertions.assertTrue(new File(filePath).exists());

        File exampleFile = new File(filePath + department + ".xlsx");
        Assertions.assertTrue(exampleFile.exists());

        targetFile = new File(filePath + "2020年年度结算确认-" + department + ".xlsx");
        if (targetFile.exists()) {
            boolean delete = targetFile.delete();
            Assertions.assertTrue(delete);
            System.out.println("Delete test file");
        }

        Files.copy(exampleFile.toPath(), targetFile.toPath());
        Assertions.assertTrue(targetFile.exists());

        ExcelTools.setFilePath(filePath);
        targetWorkbook = new XSSFWorkbook(new FileInputStream(targetFile.getAbsolutePath()));
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
     * Method: copyInnerSheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopyInnerSheet() throws IOException {
        ExcelTools.copyInnerSheet("0301.佣金特许权收入.xlsx", targetWorkbook, "广材产品部");
    }

    /**
     * Method: copySheet(String origFileName, Workbook targetWorkbook, String department) throws IOException
     */
    @Test
    public void testCopySheet() throws IOException {
        ExcelTools.copySheet("0402.专项费用预实.xlsx", targetWorkbook, "广材产品部");
        ExcelTools.copySheet("0401.人力成本预实.xlsx", targetWorkbook, "广材产品部");
        ExcelTools.copySheet("0403.基本费用预实.xlsx", targetWorkbook, "广材产品部");
    }
}
