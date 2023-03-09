package com.zyq.report.model;

import lombok.Data;

@Data
public class DistributeSaleReq {

    private Integer min;

    private Integer max;

    private String path;

    private String sheetName;

}
