package com.example.demo2;

import com.grapecity.documents.excel.*;

import java.time.LocalDate;

public class GetCacheValue {

    public static void main(String[] args) {
        IWorkbook  workbook = new Workbook();
        XlsxOpenOptions options = new XlsxOpenOptions();
        options.setDoNotRecalculateAfterOpened( true);
        workbook.open("cache.xlsx");
        workbook.setEnableCalculation( false);
        workbook.getOptions().getFormulas().setEnableIterativeCalculation(false);

        IWorksheet sheet = workbook.getWorksheets().get(0);
        Object value = sheet.getRange("A1").getValue();
        if (value instanceof CalcError) {
            System.out.println("A1: 函数错误");
        }
        System.out.println("A2: "+sheet.getRange("A2").getValue());
        System.out.println("A3: "+sheet.getRange("A3").getValue());
        System.out.println("A4: "+sheet.getRange("A4").getValue());
        System.out.println("A5: "+sheet.getRange("A5").getValue());
    }
}
