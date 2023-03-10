package com.zyq.report.controller;

import com.zyq.report.config.SingleFile;
import com.zyq.report.model.DistributeSaleReq;
import com.zyq.report.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @RequestMapping(value = "/distributeSale", method = RequestMethod.POST)
    public String distributeSale(@RequestBody DistributeSaleReq req) {
        File file = SingleFile.getFile();
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
            excelService.distributeSale(req.getMin(), req.getMax(), inputStream, req.getPath(), req.getSheetName());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "ok";
    }

}
