package com.zyq.report.model;

import lombok.Data;

@Data
public class ComputeCostReq {

    private Integer min;

    private Integer max;

    private String path;

    private String sheetName;

    private Integer costMin;

    private Integer costMax;

    private String costPath;

    private String costSheetName;


}
