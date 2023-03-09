package com.zyq.report.util;

import com.zyq.report.constant.ConmonConstant;
import com.zyq.report.model.SaleModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
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

    public List<Row> getRangeRowList(Sheet sheet, int min, int max) {
        List<Row> rows = new ArrayList<>();
        for (int i = min; i <= max; i++) {
            Row row = sheet.getRow(i);
            rows.add(row);
        }
        return rows;
    }

    public List<Row> getHeaderRowList(Sheet sheet) {
        List<Row> rows = new ArrayList<>();
        for (int i = 0; i <= 1; i++) {
            Row row = sheet.getRow(i);
            rows.add(row);
        }
        return rows;
    }

    public List<SaleModel> convertRow(List<Row> rowList) {
        List<SaleModel> saleList = new ArrayList<>();
        SaleModel saleModel = null;
        String s = "";
        for (Row row :rowList) {
            saleModel = new SaleModel();
            saleModel.setSale0(row.getCell(0) == null ? null : row.getCell(0).getStringCellValue());
            saleModel.setSale1(row.getCell(1) == null ? null : row.getCell(1).getLocalDateTimeCellValue());
            saleModel.setSale2(row.getCell(2) == null ? null : row.getCell(2).getStringCellValue());
            saleModel.setSale3(row.getCell(3) == null ? null : row.getCell(3).getStringCellValue());
            saleModel.setSale4(row.getCell(4) == null ? null : new BigDecimal(row.getCell(4).getNumericCellValue()));
            saleModel.setSale5(row.getCell(5) == null ? null : new BigDecimal(row.getCell(5).getNumericCellValue()));
            saleModel.setSale6(row.getCell(6) == null ? null : new BigDecimal(row.getCell(6).getNumericCellValue()));
            saleModel.setSale7(row.getCell(7) == null ? null : row.getCell(7).getStringCellValue());
            saleModel.setSale8(row.getCell(8) == null ? null : row.getCell(8).getStringCellValue());
            saleModel.setSale9(row.getCell(9) == null ? null : row.getCell(9).getStringCellValue());
            saleModel.setSale10(row.getCell(10) == null ? null : new BigDecimal(row.getCell(10).getNumericCellValue()));
            saleModel.setSale11(row.getCell(11) == null ? null : row.getCell(11).getStringCellValue());
            saleModel.setSale12(row.getCell(12) == null ? null : row.getCell(12).getStringCellValue());
            saleModel.setSale13(row.getCell(13) == null ? null : new BigDecimal(row.getCell(13).getNumericCellValue()));
            saleModel.setSale14(row.getCell(14) == null ? null : new BigDecimal(row.getCell(14).getNumericCellValue()));
            saleModel.setSale15(row.getCell(15) == null ? null : row.getCell(15).getCellFormula());
            saleModel.setGrossProfit(row.getCell(15) == null ? null : new BigDecimal(row.getCell(15).getNumericCellValue()));
            saleModel.setSale16(row.getCell(16) == null ? null : row.getCell(16).getStringCellValue());
            saleModel.setSale17(row.getCell(17) == null ? null : row.getCell(17).getStringCellValue());
            if (row.getCell(18) != null) {
                row.getCell(18).getCellType();
                switch (row.getCell(18).getCellType().name()) {
                    case "NUMERIC": s = String.valueOf(row.getCell(18).getNumericCellValue()); break;
                    case "STRING": s = row.getCell(18).getStringCellValue(); break;
                }
            }
            saleModel.setSale18(row.getCell(18) == null ? null : s);
            saleModel.setNum(row.getRowNum());
            saleList.add(saleModel);
        }
        return saleList;
    }

    /**
     * 设置列宽
     * @param oldSheet
     * @param newSheet
     */
    public void setColumnWidth(Sheet oldSheet, Sheet newSheet) {
        newSheet.setColumnWidth(0, oldSheet.getColumnWidth(0));
        newSheet.setColumnWidth(1, oldSheet.getColumnWidth(1));
        newSheet.setColumnWidth(2, oldSheet.getColumnWidth(2));
        newSheet.setColumnWidth(3, oldSheet.getColumnWidth(3));
        newSheet.setColumnWidth(4, oldSheet.getColumnWidth(4));
        newSheet.setColumnWidth(5, oldSheet.getColumnWidth(5));
        newSheet.setColumnWidth(6, oldSheet.getColumnWidth(6));
        newSheet.setColumnWidth(7, oldSheet.getColumnWidth(7));
        newSheet.setColumnWidth(8, oldSheet.getColumnWidth(8));
        newSheet.setColumnWidth(9, oldSheet.getColumnWidth(9));
        newSheet.setColumnWidth(10, oldSheet.getColumnWidth(10));
        newSheet.setColumnWidth(11, oldSheet.getColumnWidth(11));
        newSheet.setColumnWidth(12, oldSheet.getColumnWidth(12));
        newSheet.setColumnWidth(13, oldSheet.getColumnWidth(13));
        newSheet.setColumnWidth(14, oldSheet.getColumnWidth(14));
        newSheet.setColumnWidth(15, oldSheet.getColumnWidth(15));
        newSheet.setColumnWidth(16, oldSheet.getColumnWidth(16));
        newSheet.setColumnWidth(17, oldSheet.getColumnWidth(17));
        newSheet.setColumnWidth(18, oldSheet.getColumnWidth(18));
    }

    public void copySaller(Sheet oldSheet, Sheet newSheet, List<Row> rowList) {
        List<Row> list = new ArrayList<>();
        int i = 2;
        Row tempRow = null;
        for (Row row : rowList) {
            tempRow  = copyRow(oldSheet, newSheet, row.getRowNum(), i);
            list.add(tempRow);
            i++;
        }
    }

    /**
     * 复制表头
     * @param oldSheet
     * @param newSheet
     */
    public void copyHeader(Sheet oldSheet, Sheet newSheet) {
        List<Row> headerRowList = getHeaderRowList(oldSheet);
        List<Row> newHeaderRowList = new ArrayList<>();
        int i = 0;
        for (Row row : headerRowList) {
            Row row1 = copyRow(oldSheet, newSheet, i, i);
            if (row.getRowNum() == 0) {
                CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 1, 18);
                newSheet.addMergedRegion(cellAddresses);
            }
            newHeaderRowList.add(row1);
            i++;
        }
    }

    /**
     * 完全复制复制行
     */
    public Row copyRow(Sheet oldSheet, Sheet newSheet, int oldNum, int newNum) {
        Row oldRow = oldSheet.getRow(oldNum);
        Row newRow = newSheet.createRow(newNum);
        newRow.setHeight(oldRow.getHeight());
        newRow.setRowStyle(oldRow.getRowStyle());
        Cell oldCell = null;
        Cell newCell = null;
        for (int i = 0; i <= 18; i++) {
            oldCell = oldRow.getCell(i);
            if (oldCell == null) {
                continue;
            }
            newCell = newRow.createCell(i);
            newCell.setCellStyle(oldCell.getCellStyle());
            CellType cellType = oldCell.getCellType();
//            newCell.setCellType(cellType);
//            newCell.setCell
            switch (cellType) {
                case NUMERIC:
                    if (i == 1) {
                        newCell.setCellValue(oldCell.getLocalDateTimeCellValue());
                    }else {
                        newCell.setCellValue(oldCell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case FORMULA:
                    newCell.setCellValue(oldCell.getCellFormula());
                    break;
                case BLANK:
                    newCell.setBlank();
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                default: break;
            }
        }
        return newRow;
    }

}
