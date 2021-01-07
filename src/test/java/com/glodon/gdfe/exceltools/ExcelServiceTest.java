package com.glodon.gdfe.exceltools;

import org.junit.jupiter.api.*;
import org.springframework.boot.test.context.SpringBootTest;

/**
 * ExcelService Tester.
 *
 * @author zhangjf-a
 * @version 1.0
 * @since 01/05/2021
 */
@SpringBootTest
public class ExcelServiceTest {

    @BeforeAll
    public static void beforeAll() throws Exception {
    }

    @BeforeEach
    public void before() throws Exception {
    }

    @AfterEach
    public void after() throws Exception {
    }

    @AfterAll
    public static void afterAll() throws Exception {
    }

    /**
     * Method: readFromExcel()
     */
    @Test
    public void testReadFromExcel() throws Exception {
        //TODO: Test goes here... 
    }

    /**
     * Method: writeToExcel(List<List<Map<String, Object>>> dataList)
     */
    @Test
    public void testWriteToExcel() throws Exception {
        //TODO: Test goes here... 
    }

    /**
     * Method: parseData(List<Map<String, Object>> dataList, List<List<String>> positive, List<List<String>> negative)
     */
    @Test
    public void testParseData() throws Exception {
        //TODO: Test goes here... 
    }


    /**
     * Method: insert(Map<Integer, Set<String>> errMap, int companyID, String code)
     */
    @Test
    public void testInsert() throws Exception {
        //TODO: Test goes here... 
        /* 
        try { 
           Method method = ExcelService.getClass().getMethod("insert", Map<Integer,.class, int.class, String.class); 
           method.setAccessible(true); 
           method.invoke(<Object>, <Parameters>); 
        } catch(NoSuchMethodException e) { 
        } catch(IllegalAccessException e) { 
        } catch(InvocationTargetException e) { 
        } 
        */
    }

    /**
     * Method: formatSql(Map<Integer, Set<String>> errMap)
     */
    @Test
    public void testFormatSql() throws Exception {
        //TODO: Test goes here... 
        /* 
        try { 
           Method method = ExcelService.getClass().getMethod("formatSql", Map<Integer,.class); 
           method.setAccessible(true); 
           method.invoke(<Object>, <Parameters>); 
        } catch(NoSuchMethodException e) { 
        } catch(IllegalAccessException e) { 
        } catch(InvocationTargetException e) { 
        } 
        */
    }

} 
