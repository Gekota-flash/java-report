package com.zyq.report.controller;

import com.zyq.report.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

@RestController
@RequestMapping("/test")
public class Controller {

    @Autowired
    ExcelUtil excelUtil;

    @RequestMapping(value = "/test", method = RequestMethod.POST)
    public void test() {
        File file = new File("E:\\caiwu_pro\\新建文件夹\\2023年销售变压器-油变、干变 - 副本.xls");
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Workbook workbook = excelUtil.readExcel(inputStream);
        Sheet sheet = excelUtil.readSheet(workbook, "变压器销售总账 (2)");
        List<Row> rowList = excelUtil.getRowList(sheet);
        for (Row row : rowList) {
            System.out.println(row.getCell(0).getLocalDateTimeCellValue().toString() + row.getCell(1).toString());
        }
    }

}
