package com.zyq.report.util;

import com.zyq.report.constant.ConmonConstant;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Component
public class ExcelUtil {

    public Workbook readExcel(InputStream inputStream) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public Sheet readSheet(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName);
    }

    public List<Row> getRowList(Sheet sheet) {
        List<Row> rows = new ArrayList<>();
        int num = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < num; i++) {
            if (ConmonConstant.SALE_TOTAL_NOT_READ.contains(i)) {
                continue;
            }
            Row row = sheet.getRow(i);
            rows.add(row);
        }
        return rows;
    }

}
