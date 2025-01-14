package com.tools.doExcel.Excels;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import com.tools.doExcel.Model.ExcelData;
import com.tools.doExcel.Model.PreviousNameList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

import java.awt.Color;
import java.net.URLEncoder;

public class ExportExcelUtils {
    public static void exportExcel(HttpServletResponse response, String fileName, ExcelData data) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));
        exportExcel(data, response.getOutputStream());
    }

    public static void exportExcel(ExcelData data, OutputStream out) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            String sheetName = data.getName();
            if (null == sheetName) {
                sheetName = "Sheet1";
            }
            XSSFSheet sheet = wb.createSheet(sheetName);
            /*创建初始excel，只有一个sheet*/
            writeExcel(wb, sheet, data);
//            XSSFSheet sheet1 = wb.createSheet("previousList");
//            writeExcel(wb, sheet1, data);
            /*添加第二个sheet*/
            wirteOhterSheet(wb, "previousList");

            wb.write(out);
        } finally {
            wb.close();
        }
    }

    /*创建sheet的内容，目前只能做到写固定一个模板的，后续优化*/
    public static void createSheetData(PreviousNameList previousNameList, ExcelData data) throws IllegalAccessException {
        /*创建第二个sheet模板内容*/
        List<List<Object>> rows = new ArrayList();
        List<Object> row = new ArrayList();
        Class<? extends PreviousNameList> aClass = new PreviousNameList().getClass();
        Field[] declaredFields = aClass.getDeclaredFields();
        List<String> titles = new ArrayList<>();
        for (Field f : declaredFields) {
            titles.add(f.getName());
            f.setAccessible(true);
            row.add(f.get(previousNameList));
        }
        data.setTitles(titles);
        rows.add(row);
        data.setRows(rows);
//        data.setName("previousList");
    }

    /*写其他sheet页*/
    public static void wirteOhterSheet(XSSFWorkbook wb, String secondSheetName) throws Exception {

        /*创建第二个sheet模板内容*/
        PreviousNameList previousNameList = new PreviousNameList();
        ExcelData data = new ExcelData();
        createSheetData(previousNameList, data );
/*
        List<List<Object>> rows = new ArrayList();
        List<Object> row = new ArrayList();
        Class<? extends PreviousNameList> aClass = new PreviousNameList().getClass();
        Field[] declaredFields = aClass.getDeclaredFields();
        List<String> titles = new ArrayList<>();
        for (Field f : declaredFields) {
            titles.add(f.getName());
            f.setAccessible(true);
            row.add(f.get(reviousNameList));
        }
        data.setTitles(titles);
        rows.add(row);
        data.setRows(rows);
        data.setName("reviousNameList");
*/

        try {
            XSSFSheet sheet = wb.createSheet(secondSheetName);
            writeExcel(wb, sheet, data);

//            wb.write(out);
        } finally {
//            wb.close();
        }
    }

    private static void writeExcel(XSSFWorkbook wb, Sheet sheet, ExcelData data) {

        int rowIndex = 0;

        rowIndex = writeTitlesToExcel(wb, sheet, data.getTitles());
        writeRowsToExcel(wb, sheet, data.getRows(), rowIndex);
        autoSizeColumns(sheet, data.getTitles().size() + 1);

    }

    private static int writeTitlesToExcel(XSSFWorkbook wb, Sheet sheet, List<String> titles) {
        int rowIndex = 0;
        int colIndex = 0;

        Font titleFont = wb.createFont();
        titleFont.setFontName("simsun");
        titleFont.setBold(true);
        // titleFont.setFontHeightInPoints((short) 14);
        titleFont.setColor(IndexedColors.BLACK.index);

        XSSFCellStyle titleStyle = wb.createCellStyle();
        titleStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        titleStyle.setFillForegroundColor(new XSSFColor(new Color(182, 184, 192)));
        titleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        setBorder(titleStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));

        Row titleRow = sheet.createRow(rowIndex);
        // titleRow.setHeightInPoints(25);
        colIndex = 0;

        for (String field : titles) {
            Cell cell = titleRow.createCell(colIndex);
            cell.setCellValue(field);
            cell.setCellStyle(titleStyle);
            colIndex++;
        }

        rowIndex++;
        return rowIndex;
    }

    private static int writeRowsToExcel(XSSFWorkbook wb, Sheet sheet, List<List<Object>> rows, int rowIndex) {
        int colIndex = 0;

        Font dataFont = wb.createFont();
        dataFont.setFontName("simsun");
        // dataFont.setFontHeightInPoints((short) 14);
        dataFont.setColor(IndexedColors.BLACK.index);

        XSSFCellStyle dataStyle = wb.createCellStyle();
        dataStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        dataStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        dataStyle.setFont(dataFont);
        setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));

        for (List<Object> rowData : rows) {
            Row dataRow = sheet.createRow(rowIndex);
            // dataRow.setHeightInPoints(25);
            colIndex = 0;

            for (Object cellData : rowData) {
                Cell cell = dataRow.createCell(colIndex);
                if (cellData != null) {
                    cell.setCellValue(cellData.toString());
                } else {
                    cell.setCellValue("");
                }

                cell.setCellStyle(dataStyle);
                colIndex++;
            }
            rowIndex++;
        }
        return rowIndex;
    }

    private static void autoSizeColumns(Sheet sheet, int columnNumber) {

        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);
            sheet.autoSizeColumn(i, true);
            int newWidth = (int) (sheet.getColumnWidth(i) + 100);
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, newWidth);
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }

    private static void setBorder(XSSFCellStyle style, BorderStyle border, XSSFColor color) {
        style.setBorderTop(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
        style.setBorderBottom(border);
        style.setBorderColor(BorderSide.TOP, color);
        style.setBorderColor(BorderSide.LEFT, color);
        style.setBorderColor(BorderSide.RIGHT, color);
        style.setBorderColor(BorderSide.BOTTOM, color);
    }
}

