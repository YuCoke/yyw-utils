package com.yyw.utils;

import com.yyw.utils.properties.OfficeProperties;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

/**
 * @Description: 启动项
 * @Author: YuYaWei
 * @Date: 2024-02-04 09:43
 */
@SpringBootApplication
@EnableConfigurationProperties(OfficeProperties.class)
public class YywUtilsApplication {

    public static void main(String[] args) {
        SpringApplication.run(YywUtilsApplication.class, args);
    }

}
