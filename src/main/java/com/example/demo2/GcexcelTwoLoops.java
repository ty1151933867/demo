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
public class GcexcelTwoLoops {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    // 缓存样式对象，避免重复创建
    private static final Map<CellStyleEnum, IStyle> STYLE_CACHE = new EnumMap<>(CellStyleEnum.class);

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        int rowIndexCounter = 2;
        List<WriteExcelDto> excelDtoList = new ArrayList<>();
        Random random = new Random();

        // mock数据
        for (int i = 1; i <= 2000; i++) {
            for (int j = 1; j <= 24; j++) {
                WriteExcelDto writeExcelDto = new WriteExcelDto();
                writeExcelDto.setRowIndex(++rowIndexCounter);
                writeExcelDto.setCol1(null);
                writeExcelDto.setCol2(j);
                writeExcelDto.setCol3("10000" + i);
                writeExcelDto.setCol4("20000" + i);
                writeExcelDto.setCol5("30000" + i);
                writeExcelDto.setCol6("40000" + i);
                writeExcelDto.setCol7("50000" + i);
                writeExcelDto.setCol8("60000" + i);
                writeExcelDto.setCol9("70000" + i);
                writeExcelDto.setCol10("80000" + i);
                writeExcelDto.setCol11("90000" + i);
                writeExcelDto.setCol12("11000" + i);
                writeExcelDto.setCol13("12000" + i);
                writeExcelDto.setCol14("13000" + i);
                writeExcelDto.setCol15("14000" + i);
                writeExcelDto.setCol16("14000" + i);
                writeExcelDto.setCol17("15000" + i);
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
                    cellData.setColumnIndex(25 + k);
                    if (writeExcelDto.getCol2() == 9 || writeExcelDto.getCol2() == 16) {
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

            // 性能优化设置
            workbook.setEnableCalculation(false);
            workbook.setDeferUpdateDirtyState(true);
            workbook.setReferenceStyle(ReferenceStyle.A1);

            workbook.open(inputStream);

            IWorksheet sheet = workbook.getWorksheets().get(0);

            // 初始化样式缓存
            initializeStyleCache(workbook);

            log.info("开始生成excel，共{}行数据", excelDtoList.size());

            // 第一次循环：设置值和公式
            setValueAndFormulaLoop(sheet, excelDtoList);

            // 第二次循环：设置样式
//            setStyleLoop(sheet, excelDtoList);

            log.info("excel生成完毕");
            workbook.save("test_Gcexcel_two_loops.xlsx");

        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            long endTime = System.currentTimeMillis();
            long totalTime = endTime - startTime;
            log.info("Excel生成完成，总耗时: {} ms", totalTime);
            System.out.println("Excel生成完成，总耗时: " + totalTime + " ms (" + (totalTime / 1000.0) + " 秒)");

            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    log.warn("关闭输入流失败", e);
                }
            }
            // 清理样式缓存
            STYLE_CACHE.clear();
        }
    }

    /**
     * 初始化样式缓存 - 完整的样式定义
     */
    private static void initializeStyleCache(Workbook workbook) {
        log.info("初始化样式缓存...");

        for (CellStyleEnum styleEnum : CellStyleEnum.values()) {
            IStyle cs = workbook.getStyles().add("PSI_" + styleEnum.name());
            cs.getFont().setName("Meiryo UI");
            cs.getFont().setSize(9);
            cs.setVerticalAlignment(VerticalAlignment.Center);

            // 保持Gcexcel原有的样式配置逻辑
            configureStyleOriginal(cs, styleEnum);
            STYLE_CACHE.put(styleEnum, cs);
        }

        log.info("样式缓存初始化完成，共{}个样式", STYLE_CACHE.size());
    }

    /**
     * 保持Gcexcel原有的样式配置逻辑
     */
    private static void configureStyleOriginal(IStyle cs, CellStyleEnum styleEnum) {
        if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
        } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
        } else if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            cs.setNumberFormat("#,##0;[Red]-#,##0");
        } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
        } else if (styleEnum.getValue() == CellStyleEnum.BLANK_INDENT.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.setIndentLevel(1);
        } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.SUM_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.SUM_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING_END.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
            cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_INDENT.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.setIndentLevel(1);
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_NUM.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Right);
            cs.setNumberFormat("#,##0;[Red]-#,##0");
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_REF_SEL_STRING.getValue()) {
            cs.setHorizontalAlignment(HorizontalAlignment.Left);
            cs.getInterior().setColor(Color.FromArgb(217, 217, 217));
            cs.getInterior().setPattern(Pattern.Solid);
        }
    }

    /**
     * 第一次循环：设置值和公式
     */
    private static void setValueAndFormulaLoop(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        log.info("开始第一次循环：设置值和公式");

        for (int i = 0; i < excelDtoList.size(); i++) {
            WriteExcelDto writeExcelDto = excelDtoList.get(i);

            if (i % 1000 == 0) {
                log.info("值和公式处理进度: {}/{}", i, excelDtoList.size());
            }

            // 设置左侧基本情报的值（批量写入）
            setBasicInfoValues(sheet, writeExcelDto);

            // 设置右侧数据区域的值和公式
            setDataAreaValuesAndFormulas(sheet, writeExcelDto);
        }

        log.info("第一次循环完成：值和公式设置完毕");
    }

    /**
     * 设置基本情报的值（批量写入25列）
     */
    private static void setBasicInfoValues(IWorksheet sheet, WriteExcelDto basicData) {
        // 批量写入25列数据
        Object[][] rowData = new Object[1][25];
        rowData[0][0] = basicData.getCol1();
        rowData[0][1] = basicData.getCol2();
        rowData[0][2] = basicData.getCol3();
        rowData[0][3] = basicData.getCol4();
        rowData[0][4] = basicData.getCol5();
        rowData[0][5] = basicData.getCol6();
        rowData[0][6] = basicData.getCol7();
        rowData[0][7] = basicData.getCol8();
        rowData[0][8] = basicData.getCol9();
        rowData[0][9] = basicData.getCol10();
        rowData[0][10] = basicData.getCol11();
        rowData[0][11] = basicData.getCol12();
        rowData[0][12] = basicData.getCol13();
        rowData[0][13] = basicData.getCol14();
        rowData[0][14] = basicData.getCol15();
        rowData[0][15] = basicData.getCol16();
        rowData[0][16] = basicData.getCol17();
        rowData[0][17] = basicData.getCol18();
        rowData[0][18] = basicData.getCol19();
        rowData[0][19] = basicData.getCol20();
        rowData[0][20] = basicData.getCol21();
        rowData[0][21] = basicData.getCol22();
        rowData[0][22] = basicData.getCol23();
        rowData[0][23] = basicData.getCol24();
        rowData[0][24] = basicData.getCol25();

        IRange range = sheet.getRange(basicData.getRowIndex(), 0, 1, 25);
        range.setValue(rowData);
    }

    /**
     * 设置数据区域的值和公式
     */
    private static void setDataAreaValuesAndFormulas(IWorksheet sheet, WriteExcelDto writeExcelDto) {
        // 计算数据范围
        int startCol = Integer.MAX_VALUE;
        int endCol = Integer.MIN_VALUE;

        for (WriteExcelDto.CellData cellData : writeExcelDto.getDataMap().values()) {
            startCol = Math.min(startCol, cellData.getColumnIndex());
            endCol = Math.max(endCol, cellData.getColumnIndex());
        }

        int cols = endCol - startCol + 1;
        int rowNum = writeExcelDto.getRowIndex();

        // 准备批量数据
        Object[][] values = new Object[1][cols];
        List<int[]> formulaPositions = new ArrayList<>();

        for (WriteExcelDto.CellData cellData : writeExcelDto.getDataMap().values()) {
            int idx = cellData.getColumnIndex() - startCol;
            if (cellData.isFormula()) {
                values[0][idx] = null;
                formulaPositions.add(new int[]{cellData.getColumnIndex()});
            } else {
                values[0][idx] = cellData.getValue();
            }
        }

        // 批量写入值
        IRange dataRange = sheet.getRange(rowNum, startCol, 1, cols);
        dataRange.setValue(values);

        // 批量设置公式
        for (int[] pos : formulaPositions) {
            int colIndex = pos[0];
            WriteExcelDto.CellData cellData = writeExcelDto.getDataMap().get(colIndex - 25);
            if (cellData != null && cellData.isFormula()) {
                sheet.getRange(rowNum, colIndex).setFormula2(cellData.getFormula());
            }
        }
    }

    /**
     * 第二次循环：设置样式
     */
    private static void setStyleLoop(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        log.info("开始第二次循环：设置样式");

        for (int i = 0; i < excelDtoList.size(); i++) {
            WriteExcelDto writeExcelDto = excelDtoList.get(i);

            if (i % 1000 == 0) {
                log.info("样式处理进度: {}/{}", i, excelDtoList.size());
            }

            // 设置左侧基本情报样式
            setBasicInfoStyles(sheet, writeExcelDto, writeExcelDto.getCol2() == 24);

            // 设置右侧数据区域样式
            setDataAreaStyles(sheet, writeExcelDto);
        }

        log.info("第二次循环完成：样式设置完毕");
    }

    /**
     * 设置基本情报的样式
     */
    private static void setBasicInfoStyles(IWorksheet sheet, WriteExcelDto basicData, boolean isEnd) {
        // 设置分类列样式（第25列）
        IStyle categoryStyle = basicData.isSummaryData() ?
                STYLE_CACHE.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getSumTitleCellStyle()) :
                STYLE_CACHE.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getTitleCellStyle());
        sheet.getRange(basicData.getRowIndex(), 24).setStyle(categoryStyle);

        // 如果是结尾行，设置特殊样式
        if (isEnd) {
            applyEndRowStyles(sheet, basicData);
        }
    }

    /**
     * 应用结尾行样式
     */
    private static void applyEndRowStyles(IWorksheet sheet, WriteExcelDto basicData) {
        for (int col = 0; col <= 23; col++) {
            CellStyleEnum endStyle;
            if (basicData.isSummaryData()) {
                if (col == 1 || col == 22) {
                    endStyle = CellStyleEnum.TOTAL_BLANK_NUM_END;
                } else {
                    endStyle = CellStyleEnum.TOTAL_BLANK_STRING_END;
                }
            } else {
                if (col == 1 || col == 22) {
                    endStyle = CellStyleEnum.BLANK_NUM_END;
                } else {
                    endStyle = CellStyleEnum.BLANK_STRING_END;
                }
            }
            sheet.getRange(basicData.getRowIndex(), col).setStyle(STYLE_CACHE.get(endStyle));
        }
    }

    /**
     * 设置数据区域的样式
     */
    private static void setDataAreaStyles(IWorksheet sheet, WriteExcelDto writeExcelDto) {
        // 计算数据范围
        int startCol = Integer.MAX_VALUE;
        int endCol = Integer.MIN_VALUE;

        for (WriteExcelDto.CellData cellData : writeExcelDto.getDataMap().values()) {
            startCol = Math.min(startCol, cellData.getColumnIndex());
            endCol = Math.max(endCol, cellData.getColumnIndex());
        }

        int cols = endCol - startCol + 1;
        int rowNum = writeExcelDto.getRowIndex();

        // 批量设置样式
        IStyle dataStyle = STYLE_CACHE.get(
                TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getTitleCellStyle());
        IStyle sumDataStyle = STYLE_CACHE.get(
                TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getSumTitleCellStyle());

        IRange dataRange = sheet.getRange(rowNum, startCol, 1, cols);
        dataRange.setStyle(writeExcelDto.isSummaryData() ? sumDataStyle : dataStyle);
    }
}
