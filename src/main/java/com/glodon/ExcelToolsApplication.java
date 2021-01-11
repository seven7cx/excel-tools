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
        File exampleFile = new File(filePath + "模板.xlsx");
        List<String> departmentList = Arrays.asList("中后台-营销管理中心", "中后台-数字广联达", "中后台-EMT", "中后台-人力资源部", "中后台-审计监察部", "中后台-干部管理部", "中后台-总裁办公室", "中后台-战略管理与发展部", "中后台-法务部", "中后台-董事会办公室", "中后台-行政部", "中后台-财经管理部", "中后台-采购管理部", "中后台-产品研发管理中心", "中后台-投资管理部", "数字造价BG-云造价产品部", "数字造价BG-企业招投标产品部", "数字造价BG-国内渠道部", "数字造价BG-国际产品部", "数字造价BG-国际渠道部", "数字造价BG-广材产品部", "数字造价BG-政务渠道部", "数字造价BG-新咨询事业部", "数字造价BG-服务产品部", "数字造价BG-电子政务产品部", "数字造价BG-研发管理部", "数字造价BG-计价产品部", "数字造价BG-计量产品部", "数字造价BG-运作支持部", "数字造价BG-造价BG公共组", "数字造价BG-造价市场部", "数字施工BG-MagiCAD全球", "数字施工BG-企业管理产品部", "数字施工BG-基建产品部", "数字施工BG-建设方产品部", "数字施工BG-新建造研究院", "数字施工BG-施工BG公共组", "数字施工BG-施工产品市场部", "数字施工BG-施工战略市场部", "数字施工BG-施工渠道部", "数字施工BG-施工运营管理部", "数字施工BG-项目管理产品部", "数字施工BG-项目管理平台部", "区域平台-上海区域平台", "区域平台-山东区域平台", "区域平台-广东区域平台", "区域平台-江苏区域平台", "区域平台-河北区域平台", "区域平台-浙江区域平台", "区域平台-福建区域平台", "区域平台-重庆区域平台", "区域平台-陕西区域平台", "平台技术中心-BIMGIS平台部", "平台技术中心-CIM平台部", "平台技术中心-CTO办公室", "平台技术中心-云平台部", "平台技术中心-图形平台部", "平台技术中心-数字中台部", "平台技术中心-数据智能部", "独立BU-有巢数字BU", "独立BU-数字装修BU", "独立BU-数字金融BU", "独立BU-数字教育BU", "独立BU-数字供采BU", "独立BU-数字城市BU", "创新中心-创新中心");
        StopWatch clock = new StopWatch();
        for (String department : departmentList) {
            clock.start(department);

            File targetFile;
            if (department.contains("-")) {
                File dir = new File(filePath + department.split("-")[0] + "/");
                if (!dir.exists()) {
                    boolean mkdirs = dir.mkdirs();
                    log.info("Create dir {}, result = {}", dir.getAbsolutePath(), mkdirs);
                }

                department = department.split("-")[1];
                targetFile = new File(dir.getAbsolutePath() + "/2020年年度结算确认-" + department + ".xlsx");
            } else {
                targetFile = new File(filePath + "2020年年度结算确认-" + department + ".xlsx");
            }

            if (targetFile.exists()) {
                boolean delete = targetFile.delete();
                log.info("Delete test file {}，result = {}", department, delete);
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
        SheetCopyTools.copyIOSheet("0505.收入预实表.xlsx", targetWorkbook, department);
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
        DataFillTools.fillSummarySheet(filePath + "010102.预实利润表.xlsx", targetWorkbook, department);
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
