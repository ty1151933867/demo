package com.example.demo2;

import com.example.demo2.entity.CellStyleEnum;
import com.example.demo2.entity.TitleCategoryEnum;
import com.example.demo2.entity.WriteExcelDto;
import com.grapecity.documents.excel.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Slf4j
public class GcexcelNew {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis(); // 开始时间

        int rowIndexCounter = 2;  // 起始行
        List<WriteExcelDto> excelDtoList = new ArrayList<>();
        Random random = new Random();
        // mock数据
        for (int i = 1; i <= 2000; i++) {
            for (int j = 1; j <=24; j++) {
                WriteExcelDto writeExcelDto = new WriteExcelDto();
                writeExcelDto.setRowIndex(++rowIndexCounter);
                writeExcelDto.setCol1(null);
                writeExcelDto.setCol2(j);
                writeExcelDto.setCol3("10000"+i);
                writeExcelDto.setCol4("20000"+i);
                writeExcelDto.setCol5("30000"+i);
                writeExcelDto.setCol6("40000"+i);
                writeExcelDto.setCol7("50000"+i);
                writeExcelDto.setCol8("60000"+i);
                writeExcelDto.setCol9("70000"+i);
                writeExcelDto.setCol10("80000"+i);
                writeExcelDto.setCol11("90000"+i);
                writeExcelDto.setCol12("11000"+i);
                writeExcelDto.setCol13("12000"+i);
                writeExcelDto.setCol14("13000"+i);
                writeExcelDto.setCol15("14000"+i);
                writeExcelDto.setCol16("14000"+i);
                writeExcelDto.setCol17("15000"+i);
                writeExcelDto.setCol18(10L);
                writeExcelDto.setCol19(20L);
                writeExcelDto.setCol20(30L);
                writeExcelDto.setCol21(40L);
                writeExcelDto.setCol22(50L);
                writeExcelDto.setCol23(60L);
                writeExcelDto.setCol24(70L);
                writeExcelDto.setCol25(80L);
                Map<Integer, WriteExcelDto.CellData> dataMap = new HashMap<>();
                for (int k = 0; k < 82; k++) {
                    WriteExcelDto.CellData cellData = new WriteExcelDto.CellData();
                    cellData.setRowIndex(rowIndexCounter);
                    cellData.setColumnIndex(25+ k);
                    if (writeExcelDto.getCol2() == 9 ||  writeExcelDto.getCol2() == 16 ) {
                        cellData.setFormula("SUM(1,2,3,4,5,6,7,8,9,10)");
                    } else {
                        cellData.setValue(random.nextInt(1000));
                    }
                    dataMap.put(k, cellData);
                }
                writeExcelDto.setDataMap(dataMap);
                excelDtoList.add(writeExcelDto);
            }
        }

        InputStream inputStream = null;
        Workbook workbook = null;

        try {
            inputStream = new ClassPathResource(PSI_TEMPLATE_FILE).getInputStream();
            workbook = new Workbook();
            workbook.open(inputStream);
            workbook.setEnableCalculation(false);
            workbook.setDeferUpdateDirtyState(false);
            workbook.setReferenceStyle(ReferenceStyle.A1);
            // sheet
            IWorksheet sheet = workbook.getWorksheets().get(0);
            // style生成
            Map<CellStyleEnum, IStyle> cellStyleMap = getDataCellStyleMap(workbook);

            // 预设整列默认样式（大幅减少每行 setStyle 调用）
            if (!excelDtoList.isEmpty()) {
                int dataStartRow = excelDtoList.get(0).getRowIndex();
                int dataEndRow = excelDtoList.get(excelDtoList.size() - 1).getRowIndex();
                int totalRows = dataEndRow - dataStartRow + 1;
                applyDefaultColumnStyles(sheet, cellStyleMap, dataStartRow, totalRows);
            }

            //excel做成

            log.info("开始生成excel");
            for (int i = 0; i < excelDtoList.size(); i++) {
                WriteExcelDto writeExcelDto = excelDtoList.get(i);

                log.info("第{}行左侧基本情报写入开始", i);
                // 左侧基本情报（批量写值并只在需要时覆盖样式）
                setBasicInfoFast(sheet, writeExcelDto, cellStyleMap, writeExcelDto.getCol2() == 24);
                log.info("第{}行左侧基本情报写入结束", i);

                log.info("第{}行右侧数据写入开始", i);
                // 右侧数据：批量写入值与样式，减少单元格级别的 setValue/setStyle 调用
                // 计算 data 区间
                int startCol = Integer.MAX_VALUE;
                int endCol = Integer.MIN_VALUE;
                for (WriteExcelDto.CellData cd : writeExcelDto.getDataMap().values()) {
                    startCol = Math.min(startCol, cd.getColumnIndex());
                    endCol = Math.max(endCol, cd.getColumnIndex());
                }
                int cols = endCol - startCol + 1;

                // 准备值数组和公式标记
                Object[][] rowValues = new Object[1][cols];
                Object[][] rowFormulas = new Object[1][cols]; // 使用 Object[][] 兼容 API
                boolean allFormula = true;
                for (WriteExcelDto.CellData cd : writeExcelDto.getDataMap().values()) {
                    int idx = cd.getColumnIndex() - startCol;
                    if (cd.isFormula()) {
                        // 移除可能的外层等号，因为单元格 setFormula2 可以接受不带 '=' 的表达式
                        rowFormulas[0][idx] = cd.getFormula();
                        rowValues[0][idx] = null; // 占位
                    } else {
                        rowValues[0][idx] = cd.getValue();
                        rowFormulas[0][idx] = null;
                        allFormula = false;
                    }
                }

                // 写值（一次调用）
                IRange dataRange = sheet.getRange(writeExcelDto.getRowIndex(), startCol, 1, cols);
                dataRange.setValue(rowValues);

                // 按单元格逐个设置公式（避免将数组对象直接传入 setFormula2 导致 InvalidFormulaException）
                for (int c = 0; c < cols; c++) {
                    if (rowFormulas[0][c] != null) {
                        sheet.getRange(writeExcelDto.getRowIndex(), startCol + c).setFormula2((String) rowFormulas[0][c]);
                    }
                }

                // 批量设置样式（整个数据区相同样式）
                IStyle dataStyle = cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getTitleCellStyle());
                IStyle sumDataStyle = cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getSumTitleCellStyle());
                dataRange.setStyle(writeExcelDto.isSummaryData() ? sumDataStyle : dataStyle);

                log.info("第{}行右侧数据写入结束", i);
            }
            log.info("excel生成完毕");
            workbook.save("test_Gcexcel.xlsx");

        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            long endTime = System.currentTimeMillis(); // 结束时间
            long totalTime = endTime - startTime; // 总耗时（毫秒）
            log.info("Excel生成完成，总耗时: {} ms", totalTime);
            System.out.println("Excel生成完成，总耗时: " + totalTime + " ms (" + (totalTime / 1000.0) + " 秒)");
        }
    }


    private static void setBasicInfo(IWorksheet sheet, WriteExcelDto basicData,
                                     Map<CellStyleEnum, IStyle> cellStyleMap, boolean isEnd) {

//        1
        IRange alertFlagCell = sheet.getRange(basicData.getRowIndex(),0);
        alertFlagCell.setValue(basicData.getCol1());
        if (basicData.isSummaryData()) {
            alertFlagCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            alertFlagCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }

//        2
        IRange idCell = sheet.getRange(basicData.getRowIndex(), 1);
        idCell.setValue(basicData.getCol2());
        if (basicData.isSummaryData()) {
            idCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM));
        } else {
            idCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_NUM));
        }

        //3
        IRange sortKeyCell = sheet.getRange(basicData.getRowIndex(), 2);
        sortKeyCell.setValue(basicData.getCol3());
        if (basicData.isSummaryData()) {
            sortKeyCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            sortKeyCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }

//        4
        IRange divCell = sheet.getRange(basicData.getRowIndex(), 3);
        divCell.setValue(basicData.getCol4());
        if (basicData.isSummaryData()) {
            divCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            divCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


//        5
        IRange subCell = sheet.getRange(basicData.getRowIndex(), 4);
        subCell.setValue(basicData.getCol5());
        if (basicData.isSummaryData()) {
            subCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            subCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange consoleCell = sheet.getRange(basicData.getRowIndex(), 5);
        consoleCell.setValue(basicData.getCol6());
        if (basicData.isSummaryData()) {
            consoleCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            consoleCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange itemConsolCell = sheet.getRange(basicData.getRowIndex(), 6);
        itemConsolCell.setValue(basicData.getCol7());
        if (basicData.isSummaryData()) {
            itemConsolCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemConsolCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange itemCodeCell = sheet.getRange(basicData.getRowIndex(), 7);
        itemCodeCell.setValue(basicData.getCol8());
        if (basicData.isSummaryData()) {
            itemCodeCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemCodeCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange itemNameSymbolCell = sheet.getRange(basicData.getRowIndex(), 8);
        itemNameSymbolCell.setValue(basicData.getCol9());
        if (basicData.isSummaryData()) {
            itemNameSymbolCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemNameSymbolCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange itemNameCell = sheet.getRange(basicData.getRowIndex(), 9);
        itemNameCell.setValue(basicData.getCol10());
        if (basicData.isSummaryData()) {
            itemNameCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemNameCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange proPlantCell = sheet.getRange(basicData.getRowIndex(), 10);
        proPlantCell.setValue(basicData.getCol11());
        if (basicData.isSummaryData()) {
            proPlantCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            proPlantCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange dosageCell = sheet.getRange(basicData.getRowIndex(), 11);
        dosageCell.setValue(basicData.getCol12());
        if (basicData.isSummaryData()) {
            dosageCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            dosageCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange minInvCell = sheet.getRange(basicData.getRowIndex(), 12);
        minInvCell.setValue(basicData.getCol13());
        if (basicData.isSummaryData()) {
            minInvCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            minInvCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange maxInvCell = sheet.getRange(basicData.getRowIndex(), 13);
        maxInvCell.setValue(basicData.getCol14());
        if (basicData.isSummaryData()) {
            maxInvCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            maxInvCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange minInvQtyCell = sheet.getRange(basicData.getRowIndex(), 14);
        minInvQtyCell.setValue(basicData.getCol15());
        if (basicData.isSummaryData()) {
            minInvQtyCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            minInvQtyCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange weeklyPrdCell = sheet.getRange(basicData.getRowIndex(), 15);
        weeklyPrdCell.setValue(basicData.getCol16());
        if (basicData.isSummaryData()) {
            weeklyPrdCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            weeklyPrdCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange moqCell = sheet.getRange(basicData.getRowIndex(), 16);
        moqCell.setValue(basicData.getCol17());
        if (basicData.isSummaryData()) {
            moqCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            moqCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange lotCell = sheet.getRange(basicData.getRowIndex(), 17);
        lotCell.setValue(basicData.getCol18());
        if (basicData.isSummaryData()) {
            lotCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            lotCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange orderLt = sheet.getRange(basicData.getRowIndex(), 18);
        orderLt.setValue(basicData.getCol19());
        if (basicData.isSummaryData()) {
            orderLt.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            orderLt.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange costCell = sheet.getRange(basicData.getRowIndex(), 19);
        costCell.setValue(basicData.getCol20());
        if (basicData.isSummaryData()) {
            costCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            costCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange qtyPerPkgCell = sheet.getRange(basicData.getRowIndex(), 20);
        qtyPerPkgCell.setValue(basicData.getCol21());
        if (basicData.isSummaryData()) {
            qtyPerPkgCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            qtyPerPkgCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange isAggCell = sheet.getRange(basicData.getRowIndex(), 21);
        isAggCell.setValue(basicData.getCol22());
        if (basicData.isSummaryData()) {
            isAggCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            isAggCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        IRange sortCell = sheet.getRange(basicData.getRowIndex(), 22);
        sortCell.setValue(basicData.getCol23());
        if (basicData.isSummaryData()) {
            sortCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM));
        } else {
            sortCell.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_NUM));
        }

        IRange psiType = sheet.getRange(basicData.getRowIndex(), 23);
        psiType.setValue(basicData.getCol24());
        if (basicData.isSummaryData()) {
            psiType.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            psiType.setStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }

        IRange caterogyNameCell = sheet.getRange(basicData.getRowIndex(), 24);
        caterogyNameCell.setValue(basicData.getCol25());
        if (basicData.isSummaryData()) {
            caterogyNameCell.setStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getSumTitleCellStyle()));
        } else {
            caterogyNameCell.setStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getTitleCellStyle()));
        }
    }

    /**
     * 为数据区域整列预设默认样式，减少每行重复 setStyle 调用（只覆盖 0..23 列，列 24 由 TitleCategory 控制）
     */
    private static void applyDefaultColumnStyles(IWorksheet sheet, Map<CellStyleEnum, IStyle> cellStyleMap,
                                                 int startRow, int totalRows) {
        if (totalRows <= 0) return;

        CellStyleEnum[] defaultStylePerCol = new CellStyleEnum[24];
        // 根据原始逻辑设置每列的默认样式（非合计、非结尾）
        defaultStylePerCol[0] = CellStyleEnum.BLANK_STRING;
        defaultStylePerCol[1] = CellStyleEnum.BLANK_NUM;
        for (int c = 2; c <= 21; c++) defaultStylePerCol[c] = CellStyleEnum.BLANK_STRING;
        defaultStylePerCol[22] = CellStyleEnum.BLANK_NUM;
        defaultStylePerCol[23] = CellStyleEnum.BLANK_STRING;

        for (int col = 0; col <= 23; col++) {
            IRange colRange = sheet.getRange(startRow, col, totalRows, 1);
            colRange.setStyle(cellStyleMap.get(defaultStylePerCol[col]));
        }
    }

    /**
     * 快速写入左侧 25 列值。我们一次性写入 25 列的值，并只在需要时覆盖第 24 列（分类名称样式）和结尾行的 End 样式
     */
    private static void setBasicInfoFast(IWorksheet sheet, WriteExcelDto basicData,
                                         Map<CellStyleEnum, IStyle> cellStyleMap, boolean isEnd) {
        // 一次性写入 25 列的值
        Object[][] leftValues = new Object[1][25];
        leftValues[0][0] = basicData.getCol1();
        leftValues[0][1] = basicData.getCol2();
        leftValues[0][2] = basicData.getCol3();
        leftValues[0][3] = basicData.getCol4();
        leftValues[0][4] = basicData.getCol5();
        leftValues[0][5] = basicData.getCol6();
        leftValues[0][6] = basicData.getCol7();
        leftValues[0][7] = basicData.getCol8();
        leftValues[0][8] = basicData.getCol9();
        leftValues[0][9] = basicData.getCol10();
        leftValues[0][10] = basicData.getCol11();
        leftValues[0][11] = basicData.getCol12();
        leftValues[0][12] = basicData.getCol13();
        leftValues[0][13] = basicData.getCol14();
        leftValues[0][14] = basicData.getCol15();
        leftValues[0][15] = basicData.getCol16();
        leftValues[0][16] = basicData.getCol17();
        leftValues[0][17] = basicData.getCol18();
        leftValues[0][18] = basicData.getCol19();
        leftValues[0][19] = basicData.getCol20();
        leftValues[0][20] = basicData.getCol21();
        leftValues[0][21] = basicData.getCol22();
        leftValues[0][22] = basicData.getCol23();
        leftValues[0][23] = basicData.getCol24();
        leftValues[0][24] = basicData.getCol25();

        IRange leftRange = sheet.getRange(basicData.getRowIndex(), 0, 1, 25);
        leftRange.setValue(leftValues);

        // 第 24 列（索引 24）按类别设样式
        IStyle catStyle = basicData.isSummaryData()
                ? cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getSumTitleCellStyle())
                : cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getTitleCellStyle());
        sheet.getRange(basicData.getRowIndex(), 24).setStyle(catStyle);

        // 若为分组末尾（isEnd），覆盖需要带下边框的列样式（通常较少）
        if (isEnd) {
            for (int col = 0; col <= 23; col++) {
                CellStyleEnum endStyle;
                if (basicData.isSummaryData()) {
                    // 合计行
                    if (col == 1 || col == 22) {
                        endStyle = CellStyleEnum.TOTAL_BLANK_NUM_END;
                    } else {
                        endStyle = CellStyleEnum.TOTAL_BLANK_STRING_END;
                    }
                } else {
                    // 普通行的结尾
                    if (col == 1 || col == 22) {
                        endStyle = CellStyleEnum.BLANK_NUM_END;
                    } else {
                        endStyle = CellStyleEnum.BLANK_STRING_END;
                    }
                }
                sheet.getRange(basicData.getRowIndex(), col).setStyle(cellStyleMap.get(endStyle));
            }
            // category 列已在上面设置，无需额外处理
        }
    }


    private static Map<CellStyleEnum, IStyle> getDataCellStyleMap(Workbook workbook) {
        Map<CellStyleEnum, IStyle> dataCellStyleMap = new HashMap<>();

        for (CellStyleEnum styleEnum : CellStyleEnum.values()) {
            IStyle cs = workbook.getStyles().add("PSI_" + styleEnum.name());
            cs.getFont().setName("Meiryo UI");
            cs.getFont().setSize(9);
            cs.setVerticalAlignment(VerticalAlignment.Center);

            if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM_END.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());

                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING_END.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_INDENT.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.setIndentLevel(1);
            } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.SUM_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.SUM_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM_END.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING_END.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM_END.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());

                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING_END.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // ボーダーラインを設定
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_INDENT.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.setIndentLevel(1);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_NUM.getValue()) {
                // 右詰
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                // データ フォーマット
                cs.setNumberFormat("#,##0;[Red]-#,##0");

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_REF_SEL_STRING.getValue()) {
                // 左詰
                cs.setHorizontalAlignment(HorizontalAlignment.Left);

                // 背景色を設定
                cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
                cs.getInterior().setPattern(Pattern.Solid);
            }

            dataCellStyleMap.put(styleEnum, cs);
        }

        return dataCellStyleMap;
    }
}
