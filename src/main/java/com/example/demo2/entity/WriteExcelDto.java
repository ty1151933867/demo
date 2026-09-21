package com.example.demo2.entity;

import com.example.demo2.utils.Utils;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

@Data
public class WriteExcelDto {
    int rowIndex;

    String col1;

    Integer col2;

    private String col3;
    private String col4;

    private String col5;

    private String col6;

    private String col7;

    private String col8;
    private String col9;

    private String col10;
    private String col11;

    private String col12;

    private String col13;

    private String col14;
    private String col15;

    private String col16;
    private String col17;

    private Long col18;
    private Long col19;
    private Long col20;

    private Long col21;

    private Long col22;

    private Long col23;

    private Long col24;

    private Long col25;

    private Map<Integer, CellData> dataMap = new HashMap<>();

    public boolean isSummaryData() {
        return Objects.isNull(col23);
    }

    @Data
    @NoArgsConstructor
    public static class CellData {
        private long weekIndex;
        private LocalDate week;
        int rowIndex;
        int columnIndex;
        @JsonProperty("isFormula")
        boolean isFormula;
        String formula;
        Double formulaValue;
        Object value;


        public void setFormula(String formula) {
            this.formula = formula;
            this.isFormula = true;
        }


        public void setValue(Object value) {
            this.value = value;
            this.isFormula = false;
        }


        public void setCellAddress(int row, int column) {
            this.rowIndex = row;
            this.columnIndex = column;
        }


        public String getCellAddress() {
            return Utils.convertNumToColString(columnIndex) + (rowIndex + 1);
        }
    }

}
