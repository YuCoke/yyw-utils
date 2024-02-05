package com.yyw.utils.document;

import cn.hutool.core.bean.BeanUtil;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.yyw.utils.properties.OfficeProperties;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.Resource;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Component;

import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @Description: word工具类
 * @Author: YuYaWei
 * @Date: 2024-02-04 11:33
 */
@Component
public class WordUtils {

    /**
     * 参数标记符前缀
     */
    private static final String PARAM_PREFIX = "{{";
    /**
     * 参数标记符后缀
     */
    private static final String PARAM_SUFFIX = "}}";
    /**
     * 未选中的字符，空的框
     */
    private static final String UNSELECTED = "□";
    /**
     * 选中的字符，带钩的框
     */
    private static final String SELECTED = "☑";
    private static final Map<Boolean, String> SELECTED_MAP = new HashMap<Boolean, String>() {{
        put(true, SELECTED);
        put(false, UNSELECTED);
    }};
    private static ObjectMapper staticObjectMapper;
    private static OfficeProperties staticOfficeProperties;
    @Resource
    private ObjectMapper objectMapper;
    @Resource
    private OfficeProperties officeProperties;

    /**
     * 根据模版填充 word
     *
     * @param response 数据流
     * @param wordName 模版名
     * @param record   待填充数据
     * @param fileName 文件名
     * @Date: 2024-02-04
     */
    public static void generatorWordByTemplate(HttpServletResponse response, OfficeProperties.WordName wordName, Object record, String fileName) {
        try {
            String wordFile = staticOfficeProperties.getWordTemplatePathMap().get(wordName);
            String objectJson = staticObjectMapper.writeValueAsString(record);
            Map<String, Object> objectMap = staticObjectMapper.readValue(objectJson,
                    new TypeReference<Map<String, Object>>() {
                    });
            generatorWord(response, wordFile, objectMap, fileName);
        } catch (Exception ex) {
            throw new RuntimeException(ex);
        }
    }

    /**
     * 根据模版填充 word
     *
     * @param response     数据流
     * @param templateName 模版名
     * @param paramMap     待填充数据
     * @param fileName     文件名
     * @Date: 2024-02-04
     */
    public static void generatorWord(HttpServletResponse response, String templateName, Map<String, Object> paramMap, String fileName) {
        //获取模版
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(templateName));
            replaceWord(document, handleParamMap(paramMap));
            LocalDateTime time = LocalDateTime.now();
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + URLEncoder.encode(fileName + time + ".docx", StandardCharsets.UTF_8));
            ServletOutputStream outputStream = response.getOutputStream();
            document.write(outputStream);
        } catch (Exception ex) {
            throw new RuntimeException(ex);
        }
    }

    /**
     * 导出word
     *
     * @param outputStream 输出流
     * @param templateName 模版地址
     * @param dataMap      填充数据
     * @Date: 2024-02-04
     */
    public static void generatorWord(OutputStream outputStream, String templateName, Map<String, Object> dataMap) {
        try {
            // 读取模板文件
            XWPFDocument doc = new XWPFDocument(new FileInputStream(templateName));
            replaceWord(doc, handleParamMap(dataMap));
            // 将文档写入输出流中
            doc.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 批量压缩导出
     *
     * @param zip        压缩流
     * @param wordName   模版名称
     * @param prefixName 文件前缀
     * @param name       文件名
     * @param record     填充数据
     * @Date: 2024-02-04
     */
    public static void generatorCompressWord(ZipOutputStream zip, OfficeProperties.WordName wordName, String prefixName, String name, Object record) throws IOException {
        String wordFile = staticOfficeProperties.getWordTemplatePathMap().get(wordName);
        generatorCompressWord(zip, wordFile, prefixName, name, BeanUtil.beanToMap(record));
    }

    /**
     * 批量压缩导出
     *
     * @param zip          压缩流
     * @param templateName 模版地址
     * @param prefixName   文件前缀
     * @param name         文件名
     * @param dataMap      填充数据
     * @Date: 2024-02-04
     */
    public static void generatorCompressWord(ZipOutputStream zip, String templateName, String prefixName, String name, Map<String, Object> dataMap) throws IOException {
        LocalDateTime time = LocalDateTime.now();
        String fileName = prefixName + name + time + ".docx";
        ByteArrayOutputStream wordStream = new ByteArrayOutputStream();
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(templateName)));
        replaceWord(doc, handleParamMap(dataMap));
        // 将文档写入输出流中
        doc.write(wordStream);
        wordStream.flush();
        wordStream.close();
        byte[] wordBytes = wordStream.toByteArray();
        ByteArrayInputStream in = new ByteArrayInputStream(wordBytes);
        ZipEntry entry = new ZipEntry(fileName);
        zip.putNextEntry(entry);
        byte[] buffer = new byte[1024];
        int len;
        while ((len = in.read(buffer)) > 0) {
            zip.write(buffer, 0, len);
        }
        in.close();
        zip.closeEntry();
    }

    /**
     * 给入参map的key加上前后缀
     *
     * @param map 入参的map
     * @return 处理后的map
     */
    private static Map<String, Object> handleParamMap(Map<String, Object> map) {
        return map.entrySet().stream()
                .filter(it -> it.getValue() != null)
                .collect(
                        Collectors.toMap(it -> PARAM_PREFIX + it.getKey() + PARAM_SUFFIX, Map.Entry::getValue));
    }

    /**
     * 替换文档中的文本
     *
     * @param document word文档对象
     * @param paramMap 参数Map
     */
    private static void replaceWord(XWPFDocument document, Map<String, Object> paramMap) {
        // 替换段落中的文本
        document.getParagraphs().forEach(xwpfParagraph -> replaceParagraph(xwpfParagraph, paramMap));

        // 替换表格中的文本
        document.getTables().stream().flatMap(table -> table.getRows().stream())
                .flatMap(xwpfTableRow -> xwpfTableRow.getTableCells().stream())
                .forEach(cell -> cell.getParagraphs().forEach(
                        xwpfParagraph -> replaceParagraph(xwpfParagraph, paramMap)));
    }

    /**
     * 替换段落中的文本
     *
     * @param paragraph 段落对象
     * @param paramMap  参数Map
     */
    private static void replaceParagraph(XWPFParagraph paragraph, Map<String, Object> paramMap) {
        String fullText = paragraph.getRuns().stream().map(XWPFRun::text).collect(Collectors.joining());
        if (paramMap.keySet().stream().anyMatch(fullText::contains)) {
            for (int i = 0; i < paragraph.getRuns().size(); i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                if (i == 0) {
                    String[] nexText = {fullText};
                    paramMap.forEach((key, value) -> {
                        String stringValue = value.toString();
                        if (value instanceof Boolean) {
                            stringValue = SELECTED_MAP.get((boolean) value);
                        }
                        nexText[0] = nexText[0].replace(key, stringValue);
                    });
                    run.setText(nexText[0], 0);
                } else {
                    run.setText("", 0);
                }
            }
        }
    }

    public static Map<String, Object> paramToBoolean(Object param, int len, int num) {
        String[] array = param.toString().split(",|，");
        Map<String, Object> map = new HashMap<>();
        List<String> list = Arrays.asList(array);
        for (int i = 1; i < len + 1; i++) {
//            params[i] = list.contains(String.valueOf(i+1));
            map.put(num + "_" + i, list.contains(String.valueOf(i)));
        }
        return map;
    }

    @PostConstruct
    public void init() {
        staticObjectMapper = objectMapper;
        staticOfficeProperties = officeProperties;
    }

}
