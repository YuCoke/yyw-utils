package com.yyw.utils.document.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.converters.ReadConverterContext;
import com.alibaba.excel.converters.WriteConverterContext;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * @Description: 适用于3.x版本
 * @Author: YuYaWei
 * @Date: 2023/4/14 09:03
 */
public abstract class BaseConverter3_0<T> implements Converter<T> {
    private final Class<T> clazz;

    //	子类传入class，接收LocalDate.class,LocalDateTime .class
    public BaseConverter3_0(Class<T> clazz) {
        this.clazz = clazz;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public Class supportJavaTypeKey() {
        return clazz;
    }

    @Override
    public T convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        Class<?> type = contentProperty.getField().getType();
        if (type == BigDecimal.class) {
            String stringValue = cellData.getStringValue().replaceAll("[\\s\\p{Zs}\u00A0]", "");
            return (T) new BigDecimal(stringValue);
        }
        if (type == String.class) {
            String stringValue = cellData.getStringValue().replaceAll("[\\s\\p{Zs}\u00A0]", "");
            return (T) stringValue;
        }
        return Converter.super.convertToJavaData(cellData, contentProperty, globalConfiguration);

    }

    @Override
    public T convertToJavaData(ReadConverterContext<?> context) throws Exception {
        return Converter.super.convertToJavaData(context);
    }

    /**
     * 读取时，将Excel的值转换为Java值的规则
     */
    @Override
    public WriteCellData<?> convertToExcelData(T t, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        if (t instanceof LocalDate) {
            return new WriteCellData<>(String.format(t.toString(), "yyyy-MM-dd"));
        }
        if (t instanceof LocalTime) {
            return new WriteCellData<>(String.format(t.toString(), "HH:mm:ss"));
        }
        if (t instanceof LocalDateTime) {
            return new WriteCellData<>(String.format(t.toString(), "yyyy-MM-dd HH:mm:ss"));
        }
        if (t instanceof BigDecimal) {
            return new WriteCellData<>(BigDecimal.valueOf(Long.parseLong(t.toString())));
        }
        return new WriteCellData<>(t.toString());
    }


    @Override
    public WriteCellData<?> convertToExcelData(WriteConverterContext<T> context) {
        if (context.getValue() instanceof LocalDate) {
            return new WriteCellData<>(String.format(context.getValue().toString(), "yyyy-MM-dd"));
        }
        if (context.getValue() instanceof LocalTime) {
            return new WriteCellData<>(String.format(context.getValue().toString(), "HH:mm:ss"));
        }
        if (context.getValue() instanceof LocalDateTime) {
            return new WriteCellData<>(String.format(context.getValue().toString(), "yyyy-MM-dd HH:mm:ss"));
        }
        if (context.getValue() instanceof BigDecimal) {
            return new WriteCellData<>(BigDecimal.valueOf(Long.parseLong(context.getValue().toString())));
        }
        return new WriteCellData<>(context.getValue().toString());
    }
}
