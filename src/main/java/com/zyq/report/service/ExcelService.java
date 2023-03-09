package com.zyq.report.service;

import com.zyq.report.model.SaleModel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;

public interface ExcelService {

    File copyExcel(String path);

    void distributeSale(int min, int max, InputStream is, String path, String sheetName);

    Map<String, List<SaleModel>> typeSallers(List<SaleModel> saleList);

    Map<String, List<Row>> typeSallerRows(List<Row> rowList);

    Set<String> getSallers(List<SaleModel> saleList);

    void createSheet(Set<String> sallers, Workbook workbook);

    void setSallerHeader(Sheet sheet);

}
