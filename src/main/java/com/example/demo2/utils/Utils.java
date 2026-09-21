package com.example.demo2.utils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

@Slf4j
public class Utils {

    public static String convertNumToColString(int col) {
        int excelColNum = col + 1;
        StringBuilder colRef = new StringBuilder(2);
        int colRemain = excelColNum;

        while(colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }

            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char)(thisPart + 64);
            colRef.insert(0, colChar);
        }

        return colRef.toString();
    }

    public static Cell getCell(Sheet sheet, int rowNumber, int colNumber) {
        Row sheetRow = sheet.getRow(rowNumber);
        if (null == sheetRow) {
            sheetRow = sheet.createRow(rowNumber);
        }
        Cell cell = sheetRow.getCell(colNumber);
        if (null == cell) {
            cell = sheetRow.createCell(colNumber);
        }
        return cell;
    }

    public static void setCellValue(Cell cell, Object data) {
        if (cell == null) {
            return;
        }

        if (data instanceof Double value) {
            cell.setCellValue(value);
        } else if (data instanceof Integer value) {
            cell.setCellValue(value);
        } else if (data instanceof Float value) {
            cell.setCellValue(value);
        } else if (data instanceof Long value) {
            cell.setCellValue(value);
        } else if (data instanceof String value) {
            cell.setCellValue(value);
        }
    }

}
