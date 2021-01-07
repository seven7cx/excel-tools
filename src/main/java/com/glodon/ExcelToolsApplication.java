package com.glodon;

import com.glodon.gdfe.exceltools.StatService;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

@SpringBootApplication
public class ExcelToolsApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelToolsApplication.class, args);

        ApplicationContext context = SpringUtil.getApplicationContext();
        StatService statService = context.getBean(StatService.class);
        statService.doStat();
    }
}
