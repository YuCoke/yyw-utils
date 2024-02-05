package com.yyw.utils.document.converter;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * @author YuYawei
 * @Description: 配合序列化注解使用
 * @Test: @JsonSerialize(using = CustomDateTimeSerializer.class)
 * @Date: 2024-02-04 10:36
 */
public class CustomDateTimeSerializer extends JsonSerializer<LocalDateTime> {

    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    @Override
    public void serialize(LocalDateTime localDateTime, JsonGenerator jsonGenerator, SerializerProvider serializerProvider) throws IOException {
        String dateTimeString = localDateTime.format(FORMATTER);
        jsonGenerator.writeString(dateTimeString);
    }
}

