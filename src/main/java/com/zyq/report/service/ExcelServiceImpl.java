package com.zyq.report.service;

import com.zyq.report.model.SaleModel;
import com.zyq.report.util.ExcelUtil;
import com.zyq.report.util.LocalDateUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private ExcelUtil excelUtil;

    @Autowired
    private LocalDateUtil localDateUtil;

    @Override
    public File copyExcel(String path) {
        File source = new File(path);
        System.out.println("path:" + path);
        System.out.println("path lengh:" + path.length());
        System.out.println("isFile:" + source.isFile());
        String newPath = path.substring(0, path.lastIndexOf(".")) + localDateUtil.getDateTimeNow() + ".xlsx";
        File dest = new File(newPath);
        System.out.println("new path:" + newPath);
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
            os.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                os.close();
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return dest;
    }

    /**
     * 分发销售单
     * @param min
     * @param max
     * @param is
     * @param path
     * @param sheetName
     */
    @Override
    public void distributeSale(int min, int max, InputStream is,  String path, String sheetName) {
        Workbook workbook = excelUtil.readExcel(is);
        Sheet sheet = workbook.getSheet(sheetName);
        //获取指定范围的row
        List<Row> rowList = excelUtil.getRangeRowList(sheet, min, max);
        //转换成model
        List<SaleModel> saleModels = excelUtil.convertRow(rowList);
        //获取按销售单位分类的销售单
        Map<String, List<SaleModel>> sallersMap = typeSallers(saleModels);
        Map<String, List<Row>> sallerRowsMap = typeSallerRows(rowList);
        //获取销售单位集合
        Set<String> sallers = getSallers(saleModels);
        Iterator<String> iterator = sallers.iterator();
        String s = "";
        Sheet tempSheet = null;
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(path);
            while (iterator.hasNext()) {
                s = iterator.next();
                tempSheet = workbook.getSheet(s);
//                log.info("sheet页:{}", s);
                if (tempSheet == null) {
                    tempSheet = workbook.createSheet(s);
                    excelUtil.setColumnWidth(sheet, tempSheet);
                    excelUtil.copyHeader(sheet, tempSheet);
                    excelUtil.copySaller(sheet, tempSheet, sallerRowsMap.get(s));
                }else {
                    appendSaller(tempSheet, sallerRowsMap.get(s));
                }
            }
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 子Sheet追加销售单
     * @param tempSheet
     * @param list
     */
    private void appendSaller(Sheet tempSheet, List<Row> list) {
        if (list == null || list.size() == 0) {
            return;
        }
        List<Row> rowList = excelUtil.getRowList(tempSheet);
        List<String> existList = new ArrayList<>();
        int number = tempSheet.getPhysicalNumberOfRows();
        List<String> collect = list.stream().map(item -> item.getCell(0).getStringCellValue()).collect(Collectors.toList());
        List<String> collect1 = rowList.stream().map(item -> item.getCell(0).getStringCellValue()).collect(Collectors.toList());
        for (String s : collect1) {
            if (collect.contains(s)) {
                existList.add(s);
            }
        }
        int tempi = number;
        int j = 0;
        int f = tempi;
        for (int i = tempi; i < (tempi+list.size()); i++) {
            Cell cell = list.get(j).getCell(0);
//            log.info(cell.getStringCellValue());
            if (cell != null && !existList.contains(cell.getStringCellValue())) {
                excelUtil.copyRow(list.get(j), tempSheet, f);
                f++;
            }
            j++;
        }
    }

    /**
     * 给销售单位做映射分类
     * @param saleList
     * @return
     */
    @Override
    public Map<String, List<SaleModel>> typeSallers(List<SaleModel> saleList) {
        return saleList.stream().collect(Collectors.groupingBy(SaleModel::getSale2));
    }

    @Override
    public Map<String, List<Row>> typeSallerRows(List<Row> rowList) {
        return rowList.stream().collect(Collectors.groupingBy(row -> row.getCell(2) == null ? null : row.getCell(2).getStringCellValue()));
    }

    @Override
    public Set<String> getSallers(List<SaleModel> saleList) {
        Set<String> set = new HashSet<>();
        for (SaleModel sale : saleList) {
            set.add(sale.getSale2());
        }
        return set;
    }

    @Override
    public Map<String, BigDecimal> getCost(int min, int max, String path, String sheetName) {
        FileInputStream is = null;
        Map<String, BigDecimal> map = new HashMap<>();
        try {
            is = new FileInputStream(path);
            Workbook workbook = excelUtil.readExcel(is);
            Sheet sheet = excelUtil.readSheet(workbook, sheetName);
            List<Row> headList = excelUtil.getRangeRowList(sheet, min, min);
            Row headRow = headList.get(0);
            List<Row> list = excelUtil.getRangeRowList(sheet, min+1, max);
            int tax = 1;
            int noTax = 9;
            String tempHead = "";
            String tempNumber = "";
            for(Row row : list) {
                tax = 1;
                noTax = 9;
                for (int i = 1; i < 7; i++) {
                    tempHead = headRow.getCell(tax).getStringCellValue().toUpperCase(Locale.ROOT);
                    tempNumber = new BigDecimal(row.getCell(0).getNumericCellValue()).toString();
                    map.put("含税-" + tempHead + "-M-" + tempNumber, new BigDecimal(row.getCell(tax).getNumericCellValue()));
                    tempHead = headRow.getCell(noTax).getStringCellValue().toUpperCase(Locale.ROOT);
//                    tempNumber = new BigDecimal(row.getCell(0).getNumericCellValue()).toString();
                    map.put("不含税-" + tempHead + "-M-" + tempNumber, new BigDecimal(row.getCell(noTax).getNumericCellValue()));
                    tax++;
                    noTax++;
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return map;
    }

    /**
     * 获取成本
     * @param min
     * @param max
     * @param path
     * @param sheetName
     * @param costMin
     * @param costMax
     * @param costPath
     * @param costSheetName
     */
    @Override
    public void computeCost(int min, int max, String path, String sheetName, int costMin, int costMax, String costPath, String costSheetName) {
        Map<String, BigDecimal> costMap = getCost(costMin, costMax, costPath, costSheetName);
        Workbook workbook = null;
        FileInputStream is = null;
        try {
            is = new FileInputStream(path);
            workbook = excelUtil.readExcel(is);
            Sheet sheet = workbook.getSheet(sheetName);
            List<Row> rangeRowList = excelUtil.getRangeRowList(sheet, min, max);
            String tax = "";
            String iron = "";
            String isSpec = "";
            String type = "";
            BigDecimal num = null;
            BigDecimal cost = num;
            BigDecimal total = num;
            for (int i = 0; i < rangeRowList.size(); i++) {
                Row row = rangeRowList.get(i);
                isSpec = row.getCell(12).getStringCellValue().trim();
                if ("是".equals(isSpec)) {
                    continue;
                }
                tax = row.getCell(8).getStringCellValue().trim();
                iron = row.getCell(7).getStringCellValue().trim();
                type = row.getCell(3).getStringCellValue().trim().toUpperCase(Locale.ROOT);
                num = new BigDecimal(row.getCell(4).getNumericCellValue());
                cost = costMap.get(tax+"-全"+iron+"-" + type);
                total = num.multiply(cost);
                if (row.getCell(14) == null) {
                    row.createCell(14).setCellValue(total.toString());
                }else {
                    row.getCell(14).setCellValue(total.toString());
                }
//                row.getCell(15).setCellValue(total.toString());
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(path);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    @Override
    public void createSheet(Set<String> sallers, Workbook workbook) {
        Iterator<String> iterator = sallers.iterator();
        String s = "";
        Sheet sheet = null;
        while (iterator.hasNext()) {
            s = iterator.next();
            sheet = workbook.getSheet(s);
            if (sheet == null) {
                workbook.createSheet(s);
            }
        }
    }

    /**
     *
     * @param sheet
     */
    @Override
    public void setSallerHeader(Sheet sheet) {
        Row row1 = sheet.getRow(0);
    }

}
