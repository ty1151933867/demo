package com.example.demo2.entity;

import lombok.Getter;


@Getter
public enum CellStyleEnum {

    BLANK_NUM(1),
    BLANK_STRING(2),
    BLANK_NUM_END(3),
    BLANK_STRING_END(4),
    BLANK_INDENT(5),
    REF_SEL_NUM(6 ),
    REF_SEL_STRING(7),
    INPUT_NUM(8),
    INPUT_STRING(9),
    SUM_NUM(10),
    SUM_STRING(11),
    INPUT_NUM_END(12),
    INPUT_STRING_END(13),
    TOTAL_BLANK_NUM(14),
    TOTAL_BLANK_STRING(15),
    TOTAL_BLANK_NUM_END(16),
    TOTAL_BLANK_STRING_END(17),
    TOTAL_BLANK_INDENT(18),
    TOTAL_REF_SEL_STRING(19),
    TOTAL_SUM_NUM(20),
    TOTAL_SUM_STRING(21);


    private final int value;

    CellStyleEnum(int value) {
        this.value = value;
    }
}
