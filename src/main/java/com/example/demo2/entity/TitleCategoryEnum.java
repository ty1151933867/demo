package com.example.demo2.entity;

import java.util.Objects;

import lombok.AllArgsConstructor;
import lombok.Getter;


@Getter
@AllArgsConstructor
public enum TitleCategoryEnum {

    SALES_ACTUAL_FCST(1, 3, 0, "S",  "salesActualFcst", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    ON_FCST(2, 4, 0, "S", "onFcst", CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM),
    SS(3, 5, 0, "S",  "ss", CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM),
    INTERNATIONAL(4, 6, 0, "S",  "international", CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_INDENT, CellStyleEnum.BLANK_NUM),
    UPSWING_FCST(5, 7, 1, "S", "upswingFcst", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    DEMAND_FCST(6, 8, 1, "S", "demandFcst", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    MOVING_AVG(7, 9, 1, "S",  "movingAvg", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    SW_PY_ACTUAL(8, 10, 1, "S", "swPyActual", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    SALES_CORRECTION_REPLACEMENT(9, 11, 0, "S", "salesCorrectionReplacement", CellStyleEnum.REF_SEL_STRING, CellStyleEnum.REF_SEL_NUM, CellStyleEnum.REF_SEL_STRING, CellStyleEnum.REF_SEL_NUM),
    CONSUMP_ACTUAL_FCST(10, 12, 0, "S", "consumpActualFcst", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    CONSUMP_CORRECTION(11, 13, 0, "S",  "consumpCorrection", CellStyleEnum.INPUT_STRING, CellStyleEnum.INPUT_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    CONSUMP_CORRECTION_OTHER(12, 14, 0, "S", "consumpCorrectionOther", CellStyleEnum.INPUT_STRING, CellStyleEnum.INPUT_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    SALES_TOTAL_SIM(13, 15, 0, "S", "salesTotalSim", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    SALES_TOTAL(14, 16, 0, "S",  "salesTotal", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    PRODUCTION_ACTUAL_FCST(15, 17, 0, "P",  "productionActualFcst", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    PRO_OUT_RESULT(16, 18, 0, "P", "proOutResult", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    INBOUND_CORRECTION_REPLACEMENT(17, 19, 0, "P",  "inboundCorrectionReplacement", CellStyleEnum.INPUT_STRING, CellStyleEnum.INPUT_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    INBOUND_TOTAL(18, 20, 0, "P", "inboundTotal", CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM, CellStyleEnum.BLANK_STRING, CellStyleEnum.BLANK_NUM),
    SUPPLY_DATE_BASED(19, 21, 0, "P", "supplyDateBased", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    INVENTORY_QUANTITY_ACTUAL_SIM(20, 22, 0, "I", "inventoryQuantityActualSim", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    INVENTORY_DAYS_ACTUAL_SIM(21, 23, 0, "I",  "inventoryDaysActualSim", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    INVENTORY_QUANTITY_ACTUAL(22, 24, 0, "I",  "inventoryQuantityActual", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    INVENTORY_DAYS_ACTUAL(23, 25, 0, "I", "inventoryDaysActual", CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM, CellStyleEnum.SUM_STRING, CellStyleEnum.SUM_NUM),
    COMMENT(24, 26, 9, "I",  "comment", CellStyleEnum.INPUT_STRING_END, CellStyleEnum.INPUT_STRING_END, CellStyleEnum.BLANK_STRING_END, CellStyleEnum.BLANK_STRING_END);

    private final int seq;
    private final int rowIndex;

    private final int type;
    private final String code;
    private final String fieldName;
    private final CellStyleEnum titleCellStyle;
    private final CellStyleEnum dataCellStyle;
    private final CellStyleEnum sumTitleCellStyle;
    private final CellStyleEnum sumDataCellStyle;
    public static TitleCategoryEnum fromSeq(int seq) {
        for (TitleCategoryEnum category : TitleCategoryEnum.values()) {
            if (Objects.equals(category.getSeq(), seq)) {
                return category;
            }
        }
        return null;
    }
}
