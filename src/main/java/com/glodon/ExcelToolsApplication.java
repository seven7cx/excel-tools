package com.glodon;

import com.glodon.gfms.DataFillTools;
import com.glodon.gfms.SheetCopyTools;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.util.StopWatch;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.List;

@Slf4j
@SpringBootApplication
public class ExcelToolsApplication {

    public static void main(String[] args) throws Exception {
        SpringApplication.run(ExcelToolsApplication.class, args);

//        ApplicationContext context = SpringUtil.getApplicationContext();
//        StatService statService = context.getBean(StatService.class);
//        statService.doStat();
        log.info("Base filePath = {}", args[0]);
        createExcels(args[0]);
    }

    private static void createExcels(String filePath) throws Exception {
        File exampleFile = new File(filePath + "模板.xlsx");;
        List<String> departmentList = Arrays.asList("计价产品部", "计量产品部");
        StopWatch clock = new StopWatch();
        for (String department : departmentList) {
            clock.start(department);
            File targetFile = new File(filePath + "2020年年度结算确认-" + department + ".xlsx");
            if (targetFile.exists()) {
                boolean delete = targetFile.delete();
                log.info("Delete test file " + department);
            }

            Files.copy(exampleFile.toPath(), targetFile.toPath());

            SheetCopyTools.setFilePath(filePath);
            XSSFWorkbook targetWorkbook = new XSSFWorkbook(new FileInputStream(targetFile.getAbsolutePath()));
            SheetCopyTools.initCommonStyle(targetWorkbook);
            createExcel(targetWorkbook, department, filePath);
            after(targetFile, targetWorkbook);
            clock.stop();
        }

        log.info(clock.prettyPrint());
    }

    private static void createExcel(XSSFWorkbook targetWorkbook, String department, String filePath) throws IOException {
        log.info("[{}] 开始创建Excel", department);
        SheetCopyTools.copyIOSheet("计价产品部-收入产值.xlsx", targetWorkbook, department);
        log.info("[{}] 收入产值", department);
        SheetCopyTools.copyInnerSheet("0301.佣金特许权收入.xlsx", targetWorkbook, department);
        log.info("[{}] 佣金特许权", department);
        SheetCopyTools.copyResourceSheet("bi_resource.xlsx", targetWorkbook, department);
        log.info("[{}] 资源交易", department);
        SheetCopyTools.copyCostSheet("0402.专项费用预实.xlsx", targetWorkbook, department);
        log.info("[{}] 专项", department);
        SheetCopyTools.copyCostSheet("0401.人力成本预实.xlsx", targetWorkbook, department);
        log.info("[{}] 人力成本", department);
        SheetCopyTools.copyCostSheet("0403.基本费用预实.xlsx", targetWorkbook, department);
        log.info("[{}] 基本费用", department);
        DataFillTools.fillSummarySheet(filePath + "010102.预实利润表（销售收入.xlsx", targetWorkbook, department);
    }

    public static void after(File targetFile, XSSFWorkbook targetWorkbook) throws Exception {
        try (
            FileOutputStream fileOutputStream = new FileOutputStream(targetFile.getAbsolutePath())
        ) {
            targetWorkbook.write(fileOutputStream);
            targetWorkbook.close();
        }
    }
}
