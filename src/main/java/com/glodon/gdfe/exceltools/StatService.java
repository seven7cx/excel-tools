package com.glodon.gdfe.exceltools;

import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author zhangjf-a
 * 05.22.2020
 */
@Service
public class StatService {

    @Resource private ExcelService excelService;
    @Resource private OracleService oracleService;

    public void doStat() {
        List<String> sqlList = excelService.readFromExcel();
        List<List<Map<String, Object>>> data = new ArrayList<>(sqlList.size());
        sqlList.forEach(sql -> {
            List<Map<String, Object>> oneCompany = oracleService.getData(sql);
            data.add(oneCompany);
        });

        excelService.writeToExcel(data);
    }
}
