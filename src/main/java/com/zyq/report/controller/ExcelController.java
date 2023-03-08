package com.zyq.report.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    @RequestMapping(value = "/uploadSale", method = RequestMethod.POST)
    public String uploadSaleFile(MultipartFile file) {
        System.out.println();
        return "haha";
    }

}
