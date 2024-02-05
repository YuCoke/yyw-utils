package com.yyw.utils.document.converter;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * @author YuYawei
 * @Description: 配合反序列化注解使用
 * @Test: @JsonDeserialize(using = CustomDateTimeDeserializer.class)
 * @Date: 2024-02-04 10:36
 */
public class CustomDateTimeDeserializer extends JsonDeserializer<LocalDateTime> {

    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    @Override
    public LocalDateTime deserialize(JsonParser jsonParser, DeserializationContext deserializationContext) throws IOException {
        String dateTimeString = jsonParser.getText();

        // 如果没有带上时间，则补上当前时间
        if (!dateTimeString.contains(":")) {
            dateTimeString += " 00:00:00";
        }

        return LocalDateTime.parse(dateTimeString, FORMATTER);
    }
}

