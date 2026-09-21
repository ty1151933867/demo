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
public class GcexcelCombined {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    // 缓存样式映射，避免重复创建
    private static Map<CellStyleEnum, IStyle> CELL_STYLE_CACHE = new HashMap<>();

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        int rowIndexCounter = 2;
        List<WriteExcelDto> excelDtoList = new ArrayList<>();
        Random random = new Random();

        // mock数据 - 使用原始数据量进行测试
        for (int i = 1; i <= 2000; i++) {  // 保持与Gcexcel.java相同的200行
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
                for (int k = 0; k < 82; k++) {  // 保持与Gcexcel.java相同的82列数据
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

            // 结合两者的优化设置
            workbook.setEnableCalculation(false);
            workbook.setDeferUpdateDirtyState(true);
            workbook.setReferenceStyle(ReferenceStyle.A1);

            workbook.open(inputStream);

            IWorksheet sheet = workbook.getWorksheets().get(0);

            // 使用Gcexcel的样式创建方式，但加入缓存优化
            CELL_STYLE_CACHE = getDataCellStyleMap(workbook);

            // 预设整列默认样式（来自GcexcelNew的优化思路）
            if (!excelDtoList.isEmpty()) {
                int dataStartRow = excelDtoList.get(0).getRowIndex();
                int dataEndRow = excelDtoList.get(excelDtoList.size() - 1).getRowIndex();
                int totalRows = dataEndRow - dataStartRow + 1;
                applyDefaultColumnStyles(sheet, CELL_STYLE_CACHE, dataStartRow, totalRows);
            }

            log.info("开始生成excel，共{}行数据", excelDtoList.size());

            // 使用优化的批量处理方式
            processInBatches(sheet, excelDtoList, CELL_STYLE_CACHE);

            log.info("excel生成完毕");
            workbook.save("test_Gcexcel_combined.xlsx");

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
            // 清理缓存
            CELL_STYLE_CACHE.clear();
        }
    }

    /**
     * 批量处理数据（结合优化思路）
     */
    private static void processInBatches(IWorksheet sheet, List<WriteExcelDto> excelDtoList,
                                         Map<CellStyleEnum, IStyle> cellStyleMap) {
        int batchSize = 2000; // 适中的批处理大小
        int totalSize = excelDtoList.size();

        for (int start = 0; start < totalSize; start += batchSize) {
            int end = Math.min(start + batchSize, totalSize);
            List<WriteExcelDto> batch = excelDtoList.subList(start, end);

            log.info("处理批次 {}/{} ({}-{})",
                    (start / batchSize) + 1,
                    (totalSize + batchSize - 1) / batchSize,
                    start, end - 1);

            processBatch(sheet, batch, cellStyleMap);
        }
    }

    /**
     * 处理单个批次（优化版本）
     */
    private static void processBatch(IWorksheet sheet, List<WriteExcelDto> batch,
                                     Map<CellStyleEnum, IStyle> cellStyleMap) {
        for (int i = 0; i < batch.size(); i++) {
            WriteExcelDto writeExcelDto = batch.get(i);

            // 使用批量写入优化的左侧基本信息方法
            setBasicInfoOptimized(sheet, writeExcelDto, cellStyleMap, writeExcelDto.getCol2() == 24);

            // 右侧数据处理（优化版本）
            processDataAreaOptimized(sheet, writeExcelDto, cellStyleMap);
        }
    }

    /**
     * 优化的基本信息写入方法（结合批量写入）
     */
    private static void setBasicInfoOptimized(IWorksheet sheet, WriteExcelDto basicData,
                                              Map<CellStyleEnum, IStyle> cellStyleMap, boolean isEnd) {
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

        // 只对特殊列设置样式（减少样式设置次数）
        if (isEnd) {
            applyEndRowStyles(sheet, basicData, cellStyleMap);
        }

        // 设置分类列样式（保持原有逻辑）
        IStyle categoryStyle = basicData.isSummaryData() ?
                cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getSumTitleCellStyle()) :
                cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getTitleCellStyle());
        sheet.getRange(basicData.getRowIndex(), 24).setStyle(categoryStyle);
    }

    /**
     * 应用结尾行样式（保持原有逻辑）
     */
    private static void applyEndRowStyles(IWorksheet sheet, WriteExcelDto basicData,
                                          Map<CellStyleEnum, IStyle> cellStyleMap) {
        // 保持Gcexcel原有的样式设置逻辑
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
            sheet.getRange(basicData.getRowIndex(), col).setStyle(cellStyleMap.get(endStyle));
        }
    }

    /**
     * 优化的数据区域处理（批量处理公式和值）
     */
    private static void processDataAreaOptimized(IWorksheet sheet, WriteExcelDto writeExcelDto,
                                                 Map<CellStyleEnum, IStyle> cellStyleMap) {
        // 计算数据范围
        int startCol = Integer.MAX_VALUE;
        int endCol = Integer.MIN_VALUE;

        for (WriteExcelDto.CellData cd : writeExcelDto.getDataMap().values()) {
            startCol = Math.min(startCol, cd.getColumnIndex());
            endCol = Math.max(endCol, cd.getColumnIndex());
        }

        int cols = endCol - startCol + 1;
        int rowNum = writeExcelDto.getRowIndex();

        // 准备批量数据
        Object[][] values = new Object[1][cols];
        List<int[]> formulaPositions = new ArrayList<>();

        for (WriteExcelDto.CellData cd : writeExcelDto.getDataMap().values()) {
            int idx = cd.getColumnIndex() - startCol;
            if (cd.isFormula()) {
                values[0][idx] = null;
                formulaPositions.add(new int[]{0, idx, cd.getColumnIndex()});
            } else {
                values[0][idx] = cd.getValue();
            }
        }

        // 批量写入值
        IRange dataRange = sheet.getRange(rowNum, startCol, 1, cols);
        dataRange.setValue(values);

        // 批量设置公式
        for (int[] pos : formulaPositions) {
            int actualCol = pos[2];
            sheet.getRange(rowNum, actualCol).setFormula2(
                    writeExcelDto.getDataMap().get(actualCol - 25).getFormula());
        }

        // 批量设置样式（保持原有逻辑）
        IStyle dataStyle = cellStyleMap.get(
                TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getTitleCellStyle());
        IStyle sumDataStyle = cellStyleMap.get(
                TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getSumTitleCellStyle());

        dataRange.setStyle(writeExcelDto.isSummaryData() ? sumDataStyle : dataStyle);
    }

    /**
     * 预设默认列样式（来自GcexcelNew的优化）
     */
    private static void applyDefaultColumnStyles(IWorksheet sheet, Map<CellStyleEnum, IStyle> cellStyleMap,
                                                 int startRow, int totalRows) {
        if (totalRows <= 0) return;

        CellStyleEnum[] defaultStyles = new CellStyleEnum[24];
        defaultStyles[0] = CellStyleEnum.BLANK_STRING;
        defaultStyles[1] = CellStyleEnum.BLANK_NUM;
        for (int c = 2; c <= 21; c++) defaultStyles[c] = CellStyleEnum.BLANK_STRING;
        defaultStyles[22] = CellStyleEnum.BLANK_NUM;
        defaultStyles[23] = CellStyleEnum.BLANK_STRING;

        for (int col = 0; col <= 23; col++) {
            IRange colRange = sheet.getRange(startRow, col, totalRows, 1);
            colRange.setStyle(cellStyleMap.get(defaultStyles[col]));
        }
    }

    /**
     * 保持Gcexcel原有的样式创建方法（加入缓存优化）
     */
    private static Map<CellStyleEnum, IStyle> getDataCellStyleMap(Workbook workbook) {
        Map<CellStyleEnum, IStyle> dataCellStyleMap = new EnumMap<>(CellStyleEnum.class);

        for (CellStyleEnum styleEnum : CellStyleEnum.values()) {
            // 检查缓存
            if (CELL_STYLE_CACHE.containsKey(styleEnum)) {
                dataCellStyleMap.put(styleEnum, CELL_STYLE_CACHE.get(styleEnum));
                continue;
            }

            IStyle cs = workbook.getStyles().add("PSI_" + styleEnum.name());
            cs.getFont().setName("Meiryo UI");
            cs.getFont().setSize(9);
            cs.setVerticalAlignment(VerticalAlignment.Center);

            // 保持Gcexcel原有的样式配置逻辑
            configureStyleOriginal(cs, styleEnum);
            dataCellStyleMap.put(styleEnum, cs);
        }

        return dataCellStyleMap;
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
}
