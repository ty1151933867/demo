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
public class Gcexcel {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis(); // 开始时间

        int rowIndexCounter = 2;  // 起始行
        List<WriteExcelDto> excelDtoList = new ArrayList<>();
        Random random = new Random();
        // mock数据
        for (int i = 1; i <= 200; i++) {
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
            // sheet
            IWorksheet sheet = workbook.getWorksheets().get(0);
            // style生成
            Map<CellStyleEnum, IStyle> cellStyleMap = getDataCellStyleMap(workbook);

            //excel做成

            log.info("开始生成excel");
            for (int i = 0; i < excelDtoList.size(); i++) {
                WriteExcelDto writeExcelDto = excelDtoList.get(i);

                log.info("第{}行左侧基本情报写入开始", i);
                //左侧基本情报
                setBasicInfo(sheet,  writeExcelDto,cellStyleMap, writeExcelDto.getCol2() == 24);
                log.info("第{}行左侧基本情报写入结束", i);

                log.info("第{}行右侧数据写入开始", i);
                //右侧数据
                for (Integer key : writeExcelDto.getDataMap().keySet()) {
                    WriteExcelDto.CellData cellData = writeExcelDto.getDataMap().get(key);

                    IRange cell = sheet.getRange(cellData.getCellAddress());
                    // 有公式
                    if (cellData.isFormula()){
                        // 有公式
                        cell.setFormula2(cellData.getFormula());
                    } else {
                        cell.setValue(cellData.getValue());
                    }

                    // style设置样式
                    if (writeExcelDto.isSummaryData()) {
                        cell.setStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getSumTitleCellStyle()));
                    } else {
                        cell.setStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getTitleCellStyle()));
                    }

                }
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
