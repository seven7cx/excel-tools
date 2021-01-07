package com.glodon.gdfe.exceltools;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.stereotype.Service;

import java.text.DecimalFormat;
import java.util.*;

/**
 * @author zhangjf-a
 * 05.22.2020
 */
@Service
public class ExcelService {

    public List<String> readFromExcel() {
        Map<Integer, Set<String>> errMap = new HashMap<>();
        class TestExcelListener extends AnalysisEventListener<LinkedHashMap<Integer, String>> {
            int currentCompanyID = 0;

            @Override
            public void invoke(LinkedHashMap<Integer, String> data, AnalysisContext context) {
                if (data.get(1) != null && data.get(1).contains("公司段") && !data.get(2).contains("范围")) {
                    currentCompanyID = Integer.parseInt(data.get(2));
                } else if (data.get(0) != null && data.get(0).startsWith("5301") && !"0".equals(data.get(2))) {
                    insert(errMap, currentCompanyID, data.get(0));
                }
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {

            }
        }

        TestExcelListener listener = new TestExcelListener();
        ExcelReader excelReader = EasyExcelFactory.read("/Users/zhangjingfei/Downloads/试算表2020-01.xlsx", null, listener).build();

        ReadSheet sheet = EasyExcel.readSheet("试算表").build();
        excelReader.read(sheet);
        List<String> errSQL = formatSql(errMap);
        excelReader.finish();

        return errSQL;
    }

    private static void insert(Map<Integer, Set<String>> errMap, int companyID, String code) {
        if (!errMap.containsKey(companyID)) {
            errMap.put(companyID, new HashSet<>());
        }

        errMap.get(companyID).add(code);
    }

    private List<String> formatSql(Map<Integer, Set<String>> errMap) {
        List<String> errSQL = new ArrayList<>(errMap.size());
        final String sqlFormat = "SELECT *\n" +
            "  FROM (SELECT gcc.segment3,\n" +
            "               gcc.acc_num,\n" +
            "               gcc.acc_desc,\n" +
            "               SUM(t.begin_balance_dr) - SUM(t.begin_balance_cr) begin_amount\n" +
            "          FROM gl_balances                t,\n" +
            "               apps.cux_gl_code_combinations_v gcc\n" +
            "         WHERE t.period_name = '2020-01'\n" +
            "           AND t.code_combination_id = gcc.code_combination_id\n" +
            "           AND gcc.segment1 = '%d'\n" +
            "           AND gcc.segment3 LIKE '53%%'\n" +
            "           AND t.ledger_id = 2022\n" +
            "           AND t.actual_flag = 'A'\n" +
            "           AND gcc.summary_flag = 'N'\n" +
            "           AND gcc.enabled_flag = 'Y'\n" +
            "           AND gcc.segment3 IN (%s)\n" +
            "         GROUP BY gcc.segment3,\n" +
            "                  gcc.acc_num,\n" +
            "                  gcc.acc_desc) aaa\n" +
            " WHERE aaa.begin_amount <> 0\n" +
            " ORDER BY 1";
        errMap.forEach((k, v) -> {
            String codes = "'" + String.join("','", v) + "'";
            String sql = String.format(sqlFormat, k, codes);
            errSQL.add(sql);
        });

        return errSQL;
    }

    public void writeToExcel(List<List<Map<String, Object>>> dataList) {
        List<List<String>> positive = new ArrayList<>();
        List<List<String>> negative = new ArrayList<>();
        for (List<Map<String, Object>> companyDataList : dataList) {
            parseData(companyDataList, positive, negative);
        }

//        ExcelWriter excelWriter = EasyExcel.write("/Users/zhangjingfei/Downloads/凭证.xlsx").build();
//        WriteSheet positiveSheet = EasyExcel.writerSheet(0, "正的").build();
//        excelWriter.write(positive, positiveSheet);
//
//        WriteSheet negativeSheet = EasyExcel.writerSheet(1, "负的").build();
//        excelWriter.write(negative, negativeSheet);
//
//        excelWriter.finish();
    }

    public void parseData(List<Map<String, Object>> dataList, List<List<String>> positive, List<List<String>> negative) {
        String companyID = "dd";
        DecimalFormat df = new DecimalFormat("#.00");
        List<List<String>> sheetData = new ArrayList<>(dataList.size());
        double total = 0;
        for (Map<String, Object> data : dataList) {
            List<String> line = new ArrayList<>(12);
            List<String> split = Arrays.asList(data.get("ACC_NUM").toString().split("\\."));
            companyID = split.get(0);
            line.addAll(split);
            line.add(data.get("BEGIN_AMOUNT").toString());
            line.add("0");
            sheetData.add(line);

            total += Double.parseDouble(data.get("BEGIN_AMOUNT").toString());
        }

        String amount = df.format(total);
        sheetData.forEach(line -> line.set(11, amount));
        positive.addAll(sheetData);

        total = 0;
        List<List<String>> sheetData2 = new ArrayList<>(dataList.size());
        for (Map<String, Object> data : dataList) {
            List<String> line = new ArrayList<>(12);
            List<String> split = Arrays.asList(data.get("ACC_NUM").toString().split("\\."));
            line.addAll(split);
            line.add(String.valueOf(-1.0 * Double.parseDouble(data.get("BEGIN_AMOUNT").toString())));
            line.add("0");
            sheetData2.add(line);

            total += Double.parseDouble(data.get("BEGIN_AMOUNT").toString()) * -1;
        }

        String negativeAmount = df.format(total);
        sheetData2.forEach(line -> line.set(11, negativeAmount));
        negative.addAll(sheetData2);

        ExcelWriter excelWriter = EasyExcel.write("/Users/zhangjingfei/Downloads/temp/" + companyID + ".xlsx").build();
        WriteSheet positiveSheet = EasyExcel.writerSheet(0, "正的").build();
        excelWriter.write(sheetData, positiveSheet);

        WriteSheet negativeSheet = EasyExcel.writerSheet(1, "负的").build();
        excelWriter.write(sheetData2, negativeSheet);

        excelWriter.finish();
    }
}
