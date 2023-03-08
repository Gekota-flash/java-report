package com.zyq.report.controller;

import com.zyq.report.config.SingleFile;
import com.zyq.report.model.PathReq;
import com.zyq.report.service.ExcelService;
import com.zyq.report.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.util.List;

@RestController
@RequestMapping("/common")
@Slf4j
public class Controller {

    @Autowired
    ExcelUtil excelUtil;

    @Autowired
    ExcelService excelService;

    @RequestMapping(value = "/init", method = RequestMethod.POST)
    public void init(@RequestBody PathReq req) {
        log.info("初始化路径:{}",req.getPath());
        File file = excelService.copyExcel(req.getPath());
        SingleFile.setFile(file);
    }

    @RequestMapping(value = "/flushPath", method = RequestMethod.POST)
    public String flushPath() {
        File file = SingleFile.getFile();
        if (file == null) {
            return "";
        }
        return file.getPath();
    }

    @RequestMapping(value = "/setPath", method = RequestMethod.POST)
    public void setPath(@RequestBody PathReq req) {
        log.info("设置路径:{}",req.getPath());
        File file = new File(req.getPath());
        SingleFile.setFile(file);
    }

    @RequestMapping(value = "/test", method = RequestMethod.POST)
    public void test() {
        File file = SingleFile.getFile();
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
            System.out.println(row.getCell(0).toString() + row.getCell(1).getLocalDateTimeCellValue().toString() + row.getCell(2).toString());
        }
//        excelService.copyExcel("E:\\caiwu_pro\\新建文件夹\\2023年销售变压器-油变、干变 - 副本1.xlsx");
        try {
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
