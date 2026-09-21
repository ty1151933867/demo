package com.example.demo2;

import com.grapecity.documents.excel.*;

import java.time.LocalDate;

public class CacheValue {

    public static void main(String[] args) {
        IWorkbook  workbook = new Workbook();
        workbook.setEnableCalculation( false);
        IWorksheet sheet = workbook.getWorksheets().get(0);
        sheet.getRange("A1").setFormula("=SUM(10,20)");
        sheet.getRange("A2").setValue(LocalDate.of(2026, 5, 19));
        sheet.getRange("A3").setValue(null);
        sheet.getRange("A4").setValue("");
        sheet.getRange("A5").setValue(1000);
        XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();
        xlsxSaveOptions.setIsCompactMode(false);
        workbook.save("cache.xlsx", xlsxSaveOptions);
    }
}
