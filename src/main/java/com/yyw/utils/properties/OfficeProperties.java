package com.yyw.utils.properties;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;

import java.util.Map;

/**
 * @Description: 配置导出模板位置和名称 配合配置类注解使用
 * @Test: @EnableConfigurationProperties(OfficeProperties.class)
 * @Author: YuYaWei
 * @Date: 2024-02-04 10:22
 */

@Data
@ConfigurationProperties(prefix = "office")
public class OfficeProperties {

    /**
     * word文档模板的路径
     */
    private Map<WordName, String> wordTemplatePathMap;
    private Map<ExcelName, String> excelTemplatePathMap;

    public enum WordName {
        /**
         * 自定义模版名称配置
         */
        WORD_NAME_1,
        WORD_NAME_2,

    }

    public enum ExcelName {
        /**
         * 自定义模版名称配置
         */
        EXCEL_NAME_1,
        EXCEL_NAME_2,
    }

}
