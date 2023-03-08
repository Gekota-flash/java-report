package com.zyq.report.util;

import org.springframework.stereotype.Component;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Component
public class LocalDateUtil {

    private DateTimeFormatter COMMON_FORMAT = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

    /**
     * 获取当前时间
     * @return yyyyMMddHHmmss
     */
    public String getDateTimeNow() {
        return LocalDateTime.now().format(COMMON_FORMAT);
    }

}
