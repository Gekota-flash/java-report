package com.zyq.report.config;

import org.springframework.stereotype.Component;

import java.io.File;

@Component
public class SingleFile {

    private SingleFile() {}

    private static File FILE = null;

    public static File getFile() {
        return FILE;
    }

    public static void setFile(File file) {
        FILE = file;
    }

}
