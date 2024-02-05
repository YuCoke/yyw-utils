package com.yyw.utils.document.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;

/**
 * @Description: 适用于2.x版本
 * @Author: YuYaWei
 * @Date: 2023/4/14 09:03
 */
public abstract class BaseConverter2<T> implements Converter<T> {
    private Class<T> clazz;

    //	子类传入class，接收LocalDate.class,LocalDateTime .class
    public BaseConverter2(Class<T> clazz) {
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

//	@Override
//	public T convertToJavaData(CellData cellData, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
//		// LocalDate时间转换
//		if (cellData.getData() instanceof LocalDate) {
//			if (cellData.getType().equals(CellDataTypeEnum.NUMBER)) {
//				LocalDate localDate = LocalDate.of(1900, 1, 1);
//				localDate = localDate.plusDays(cellData.getNumberValue().longValue());
//				return (T) localDate;
//			} else if (cellData.getType().equals(CellDataTypeEnum.STRING)) {
//				return (T) LocalDate.parse(cellData.getStringValue(), DateTimeFormatter.ofPattern("yyyy-MM-dd"));
//			} else {
//				return null;
//			}
//		}
//
//		// LocalDateTime 时间转换
//		if (cellData.getData() instanceof LocalDateTime) {
//			return (T) LocalDateTime.parse(cellData.getStringValue(), DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
//		}
//		return null;
//	}
//
//	@Override
//	public CellData convertToExcelData(T t, ExcelContentProperty excelContentProperty, GlobalConfiguration globalConfiguration) throws Exception {
//		if (t instanceof LocalDate) {
//			return new CellData<>(String.format(t.toString(), "yyyy-MM-dd"));
//		}
//
//		if (t instanceof LocalDateTime) {
//			return new CellData<>(String.format(t.toString(),"yyyy-MM-dd HH:mm:ss"));
//		}
//		return new CellData<>(t.toString());
//	}
}
