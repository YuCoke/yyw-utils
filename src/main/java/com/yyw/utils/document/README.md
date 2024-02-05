# 使用说明

## 一、Excel使用操作

参考[EasyExcel官方地址](https://easyexcel.opensource.alibaba.com/)

1. 导出
    ```java
     HttpServletResponse response;
     List<T> data = new ArrayList<>(); //替换为需要导出的数据
     ExcelUtils.generatorExcel(response, OfficeProperties.ExcelName.Excel_Name, "fileName", data);
    
    ```

2. 导入 -- **合并单元格自动拆分的导入**
    ```java
     MultipartFile file; //上传的文件
     int success = 0;
     int error = 0;
     ImportExcelListener<F> listener = new ImportExcelListener<>(1);
     List<F> fields = listener.getData(file.getInputStream(), F.class);
     for (F field : fields) {
         T entity = fieldMapper.toEntity(field);
         try {
             TService.save(entity);
             ++success;
         } catch (Exception e) {
             Throwable cause = e.getCause();
             if (cause instanceof SQLException) {
                 SQLException sql = (SQLException) cause;
             }
             ++error;
         }
     }
    ```

## 二、Word使用操作

1. 导出
    ```java
     HttpServletResponse response;
     F record = new F();
     WordUtils.generatorWord(response, OfficeProperties.WordName.WORD_NAME, record, fileName);
    ```
2. 批量导出
    ```java
     List<T> records = new ArrayList<>();
     // 批量生成 Word 文档
     ByteArrayOutputStream zipStream = new ByteArrayOutputStream();
     ZipOutputStream zip = new ZipOutputStream(zipStream);
     for (T record : records) {
         WordUtils.generatorCompressWord(zip, OfficeProperties.WordName.FIELD_INSPECTION_RECORDE, "prefixName", fileName, record);
     }
     zip.close();

     // 下载压缩文件
     response.setContentType("application/octet-stream");
     LocalDateTime today = LocalDateTime.now();
     response.setHeader("Content-Disposition",
                        "attachment;filename=" + URLEncoder.encode("zipName-" + today + ".zip", "UTF-8"));
     response.getOutputStream().write(zipStream.toByteArray());
    ```
   
