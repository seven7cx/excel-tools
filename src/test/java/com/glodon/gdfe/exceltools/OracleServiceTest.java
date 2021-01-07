package com.glodon.gdfe.exceltools;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.util.StringUtils;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * OracleService Tester.
 *
 * @author zhangjf-a
 * @version 1.0
 * @since 05/22/2020
 */
@SpringBootTest
@RunWith(SpringRunner.class)
public class OracleServiceTest {

    @Autowired private OracleService oracleService;

    @Before
    public void before() throws Exception {
    }

    @After
    public void after() throws Exception {
    }

    /**
     * Method: getData(String sql)
     */
    @Test
    public void testGetData() throws Exception {
        String sql = "SELECT *\n" +
            "  FROM (SELECT gcc.segment3,\n" +
            "               gcc.acc_num,\n" +
            "               gcc.acc_desc,\n" +
            "               SUM(t.begin_balance_dr) - SUM(t.begin_balance_cr) begin_amount\n" +
            "          FROM gl_balances                t,\n" +
            "               cux_gl_code_combinations_v gcc\n" +
            "         WHERE t.period_name = '2020-01'\n" +
            "           AND t.code_combination_id = gcc.code_combination_id\n" +
            "           AND gcc.segment1 = '160'\n" +
            "           AND gcc.segment3 LIKE '53%'\n" +
            "           AND t.ledger_id = 2022\n" +
            "           AND t.actual_flag = 'A'\n" +
            "           AND gcc.summary_flag = 'N'\n" +
            "           AND gcc.enabled_flag = 'Y'\n" +
            "           AND gcc.segment3 IN ('53013101','53014022','53013308','53013109','53013108','53013107','53013106','53013304','53013105','53013104','53014028','53014005')\n" +
            "         GROUP BY gcc.segment3,\n" +
            "                  gcc.acc_num,\n" +
            "                  gcc.acc_desc) aaa\n" +
            "WHERE aaa.begin_amount <> 0\n" +
            "ORDER BY 1\n";
        List<Map<String, Object>> data = oracleService.getData(sql);
        Assert.assertFalse(StringUtils.isEmpty(data));
        System.out.println(data.get(0));
    }

    @Test
    public void testReadSVN() throws IOException {
        Map<String, Map<String, String>> svn = new HashMap<>();
        String currentRepo = null;
        try (
            BufferedReader userReader = new BufferedReader(new FileReader("E:/test/svn项目权限.sql"))
        ) {
            String line;
            while ((line = userReader.readLine()) != null) {
                if (line.startsWith("[repos1:")) {
                    currentRepo = line;
                    if (!svn.containsKey(currentRepo)) {
                        svn.put(currentRepo, new HashMap<>());
                    }
                } else if (line.endsWith("]")) {
                    currentRepo = null;
                } else {
                    if (currentRepo != null && line.contains("=")) {
                        String[] authority = line.split("=");
                        Map<String, String> authorityMap = svn.get(currentRepo);
                        if (authorityMap != null) {
                            authorityMap.put(authority[0].trim(), authority[1].trim());
                        }
                    }
                }
            }
        }

        svn.forEach((k, map) -> {
            System.out.println(k);
            System.out.println(map);
        });
    }

} 
