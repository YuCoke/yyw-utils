package com.yyw.utils.document;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import org.apache.commons.collections4.CollectionUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 多sheet、有合并单元格的excel导入
 *
 * @author Yuyawei
 */
public class ImportExcelListener<T> extends AnalysisEventListener<T> {

    private static final Logger LOGGER = LoggerFactory.getLogger(ImportExcelListener.class);
    /**
     * 最终返回的解析数据list
     */
    private final List<T> data = new ArrayList<>();
    /**
     * 解析数据
     * key是sheetName，value是相应sheet的解析数据
     */
    private final Map<String, List<T>> dataMap = new HashMap<>();
    /**
     * 合并单元格
     * key键是sheetName，value是相应sheet的合并单元格数据
     */
    private final Map<String, List<CellExtra>> mergeMap = new HashMap<>();
    /**
     * 正文起始行
     */
    private final Integer headRowNumber;

    public ImportExcelListener(Integer headRowNumber) {
        this.headRowNumber = headRowNumber;
    }

    @Override
    public void invoke(T data, AnalysisContext context) {
        String sheetName = context.readSheetHolder().getSheetName();
        dataMap.computeIfAbsent(sheetName, k -> new ArrayList<>());
        dataMap.get(sheetName).add(data);
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        String sheetName = context.readSheetHolder().getSheetName();
        switch (extra.getType()) {
            // 额外信息是合并单元格
            case MERGE:
                if (extra.getRowIndex() >= headRowNumber) {
                    mergeMap.computeIfAbsent(sheetName, k -> new ArrayList<>());
                    mergeMap.get(sheetName).add(extra);
                }
                break;
            case COMMENT:
            case HYPERLINK:
            default:
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        LOGGER.info("Excel解析完成");
    }

    /**
     * 获取解析数据
     */
    public List<T> getData(InputStream in, Class<T> clazz) {
        try {
            EasyExcel.read(in, clazz, this)
                    .extraRead(CellExtraTypeEnum.MERGE)
                    .headRowNumber(headRowNumber)
                    .doReadAll();
        } catch (Exception e) {
            LOGGER.error("Excel读取异常：" + e);
            throw new RuntimeException("读取失败，请确认上传文件是否符合该导入模版并检查字段是否符合该字段要求的类型");
        }
        convertDataMapToData();
        return data;
    }

    /**
     * 将具有多个sheet数据的dataMap转变成一个data
     */
    private void convertDataMapToData() {
//        Iterator<Map.Entry<String, List<T>>> iterator = dataMap.entrySet().iterator();
//        while (iterator.hasNext()) {
//        Map.Entry<String, List<T>> next = iterator.next();

        for (Map.Entry<String, List<T>> next : dataMap.entrySet()) {
            String sheetName = next.getKey();
            List<T> list = next.getValue();
            List<CellExtra> mergeList = mergeMap.get(sheetName);
            if (CollectionUtils.isNotEmpty(mergeList)) {
                explainMergeData(list, mergeList);
            }
            data.addAll(list);
        }
    }

    /**
     * 处理有合并单元格的数据
     *
     * @param list      解析数据
     * @param mergeList 合并单元格信息
     * @return 填充好的解析数据
     */
    private List<T> explainMergeData(List<T> list, List<CellExtra> mergeList) {
        // 循环所有合并单元格信息
        mergeList.forEach(item -> {
            Integer firstRowIndex = item.getFirstRowIndex() - headRowNumber;
            int lastRowIndex = item.getLastRowIndex() - headRowNumber;
            Integer firstColumnIndex = item.getFirstColumnIndex();
            Integer lastColumnIndex = item.getLastColumnIndex();
            // 获取初始值
            Object initValue = getInitValueFromList(firstRowIndex, firstColumnIndex, list);
            // 设置值
            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                    setInitValueToList(initValue, i, j, list);
                }
            }
        });
        return list;
    }

    /**
     * 获取合并单元格的初始值
     * rowIndex对应list的索引
     * columnIndex对应实体内的字段
     *
     * @param firstRowIndex    起始行
     * @param firstColumnIndex 起始列
     * @param list             列数据
     * @return 初始值
     */
    private Object getInitValueFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> list) {
        Object filedValue = null;
        T object = list.get(firstRowIndex);
        for (Field field : object.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == firstColumnIndex) {
                    try {
                        filedValue = field.get(object);
                        break;
                    } catch (IllegalAccessException e) {
                        LOGGER.error("获取合并单元格的初始值异常：" + e.getMessage());
                    }
                }
            }
        }
        return filedValue;
    }

    /**
     * 设置合并单元格的值
     *
     * @param filedValue  值
     * @param rowIndex    行
     * @param columnIndex 列
     * @param list        解析数据
     */
    public void setInitValueToList(Object filedValue, Integer rowIndex, Integer columnIndex, List<T> list) {
        T object = list.get(rowIndex);
        for (Field field : object.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == columnIndex) {
                    try {
                        field.set(object, filedValue);
                        break;
                    } catch (IllegalAccessException e) {
                        LOGGER.error("设置合并单元格的值异常：" + e.getMessage());
                    }
                }
            }
        }
    }
}
