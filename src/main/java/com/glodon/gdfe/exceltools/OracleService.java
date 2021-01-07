package com.glodon.gdfe.exceltools;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import javax.annotation.Resource;
import java.util.List;
import java.util.Map;

/**
 * @author zhangjf-a
 * 05.22.2020
 */
@Repository
public class OracleService {

    @Resource private JdbcTemplate jdbcTemplate;

    public List<Map<String, Object>> getData(String sql) {
        return jdbcTemplate.queryForList(sql);
    }
}
