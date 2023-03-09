package com.zyq.report.controller;

import com.zyq.report.config.SingleFile;
import com.zyq.report.model.PathReq;
import com.zyq.report.service.ExcelService;
import com.zyq.report.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.*;

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
//        File file1 = SingleFile.getFile();
        InputStream inputStream = null;
//        OutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(file);
//            outputStream = new FileOutputStream(file1);

        excelService.distributeSale(2,81, inputStream, file.getPath(),"变压器销售总账 (2)");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
//                outputStream.close();
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
