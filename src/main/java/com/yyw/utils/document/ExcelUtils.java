package com.yyw.utils.document;

import cn.hutool.core.bean.BeanUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.yyw.utils.document.converter.LocalDateConverter;
import com.yyw.utils.properties.OfficeProperties;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.Resource;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author YuYaWei
 * @Description: excel工具类
 * @createTime 2023-04-13 14:16
 */
@Component
public class ExcelUtils {

    public static final String ENTITY_ERROR = "导出数据不能为空";
    private static OfficeProperties staticOfficeProperties;
    @Resource
    private OfficeProperties officeProperties;

    /**
     * 根据模版导出
     *
     * @param response     数据流
     * @param templateName 模版名
     * @param outFileName  文件名
     * @param bean         导出数据 不分类型
     * @Date: 2024-02-04
     */
    public static void generatorExcelByTemplate(HttpServletResponse response, OfficeProperties.ExcelName templateName, String outFileName, Object bean) {
        String excelFile = staticOfficeProperties.getExcelTemplatePathMap().get(templateName);
        List<Map<String, Object>> map = null;
        if (bean instanceof List) {
            map = ((List<?>) bean).stream().map(BeanUtil::beanToMap).collect(Collectors.toList());
        } else if (bean != null) {
            map = Collections.singletonList(BeanUtil.beanToMap(bean));
        }
        easyExcelExport(response, excelFile, outFileName, map);
    }

    /**
     * 根据模版导出 列表填充
     *
     * @param response         数据流
     * @param templateFileName 模版名
     * @param outFileName      文件名
     * @param list             导出数据
     * @Date: 2024-02-04
     */
    public static void easyExcelExport(HttpServletResponse response, String templateFileName, String outFileName, List<?> list) {
        if (list == null || list.isEmpty()) {
//            throw new ServiceException(ENTITY_ERROR);
            throw new RuntimeException(ENTITY_ERROR);
        }
        if (templateFileName == null || templateFileName.isEmpty()) {
            throw new RuntimeException("excel模板文件路径不能为空");
        }
        ExcelWriter excelWriter = null;
        try {
            String encodeName = outFileName + LocalDateTime.now();
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + URLEncoder.encode(encodeName + ".xlsx", StandardCharsets.UTF_8));
            LocalDateConverter localDateConverter = new LocalDateConverter();
            excelWriter = EasyExcel.write(response.getOutputStream()).registerConverter(localDateConverter).withTemplate(templateFileName).build();
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            //填充数据以外的内容在这以前先处理
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
            List data = new ArrayList();
            Integer index = 1;
            for (Object o : list) {
                Map<String, Object> map = BeanUtil.beanToMap(o);
                map.put("index", index);
                data.add(map);
                index++;
            }
            excelWriter.fill(data, fillConfig, writeSheet);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            // 关闭流
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 直接根据EasyExcel配置导出
     *
     * @param response 数据流
     * @param fileName 文件名
     * @param list     导出数据
     * @Date: 2024-02-04
     */
    public static void generatorExcel(HttpServletResponse response, String fileName, List<T> list) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("UTF-8");
        String name = URLEncoder.encode(fileName + LocalDateTime.now(), StandardCharsets.UTF_8);
        response.setHeader("Content-disposition", "attachment;filename=" + name + ".xlsx");
        EasyExcel.write(response.getOutputStream(), T.class)
                .sheet()
                .head(T.class)
                .doWrite(list);
    }

    @PostConstruct
    public void init() {
        staticOfficeProperties = officeProperties;
    }
}



