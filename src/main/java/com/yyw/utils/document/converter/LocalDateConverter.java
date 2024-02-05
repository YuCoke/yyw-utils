package com.yyw.utils.document.converter;

import java.time.LocalDate;

/**
 * @Description: 配合EasyExcel注解使用
 * @Test: @ExcelProperty(value = "xxx",converter = LocalDateTimeConverter.class)
 * @Author: YuYaWei
 * @Date: 2024-02-04 10:36
 */

public class LocalDateConverter extends BaseConverter3_0<LocalDate> {
    public LocalDateConverter() {
        super(LocalDate.class);
    }
}
