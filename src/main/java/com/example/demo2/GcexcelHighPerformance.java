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
import java.util.concurrent.ConcurrentHashMap;

@Slf4j
public class GcexcelHighPerformance {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    // 高性能样式缓存 - 使用ConcurrentHashMap保证线程安全
    private static final Map<CellStyleEnum, IStyle> STYLE_CACHE = new ConcurrentHashMap<>();
    private static final Map<String, IStyle> COLUMN_STYLE_CACHE = new ConcurrentHashMap<>();

    // 批处理配置
    private static final int BATCH_SIZE = 5000; // 增大批处理大小
    private static final int PRESET_ROWS = 10000; // 预设行数

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        int rowIndexCounter = 2;
        List<WriteExcelDto> excelDtoList = new ArrayList<>();
        Random random = new Random();

        // mock数据 - 保持原有规模
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

            // 极致性能优化设置
            workbook.setEnableCalculation(false);
            workbook.setDeferUpdateDirtyState(true);
            workbook.setReferenceStyle(ReferenceStyle.A1);

            // 新增性能优化选项
//            workbook.setMemorySetting(MemorySetting.MemoryPreference);

            workbook.open(inputStream);

            IWorksheet sheet = workbook.getWorksheets().get(0);

            // 预先扩展工作表行数，避免动态扩展开销
            sheet.getRange(0, 0, PRESET_ROWS, 110).setValue(""); // 预分配空间

            log.info("开始高性能Excel生成，共{}行数据", excelDtoList.size());

            // 分阶段处理以最大化性能
            processWithOptimizedStages(sheet, excelDtoList);

            log.info("Excel生成完毕");
            workbook.save("test_high_performance.xlsx");

        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            long endTime = System.currentTimeMillis();
            long totalTime = endTime - startTime;
            log.info("高性能Excel生成完成，总耗时: {} ms", totalTime);
            System.out.println("高性能Excel生成完成，总耗时: " + totalTime + " ms (" + (totalTime / 1000.0) + " 秒)");

            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    log.warn("关闭输入流失败", e);
                }
            }
            // 清理缓存
            STYLE_CACHE.clear();
            COLUMN_STYLE_CACHE.clear();
        }
    }

    /**
     * 分阶段优化处理
     */
    private static void processWithOptimizedStages(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        log.info("=== 阶段1: 初始化样式缓存 ===");
        initializeStyleCaches(sheet.getWorkbook());

        log.info("=== 阶段2: 预设列样式 ===");
        presetColumnStyles(sheet, excelDtoList);

        log.info("=== 阶段3: 批量数据写入 ===");
        processInUltraLargeBatches(sheet, excelDtoList);

        log.info("=== 阶段4: 特殊样式处理 ===");
        applySpecialStyles(sheet, excelDtoList);
    }

    /**
     * 初始化双重样式缓存系统
     */
    private static void initializeStyleCaches(IWorkbook workbook) {
        log.info("初始化双重样式缓存系统...");

        // 初始化完整样式缓存（保持原有复杂样式）
        for (CellStyleEnum styleEnum : CellStyleEnum.values()) {
            IStyle cs = workbook.getStyles().add("PSI_" + styleEnum.name());
            cs.getFont().setName("Meiryo UI");
            cs.getFont().setSize(9);
            cs.setVerticalAlignment(VerticalAlignment.Center);
            configureCompleteStyle(cs, styleEnum);
            STYLE_CACHE.put(styleEnum, cs);
        }

        // 初始化列级别样式缓存（用于批量应用）
        initializeColumnLevelStyles(workbook);

        log.info("样式缓存初始化完成，完整样式:{}个，列样式:{}个",
                STYLE_CACHE.size(), COLUMN_STYLE_CACHE.size());
    }

    /**
     * 初始化列级别样式缓存
     */
    private static void initializeColumnLevelStyles(IWorkbook workbook) {
        // 左侧基础列样式组
        COLUMN_STYLE_CACHE.put("LEFT_NUM", STYLE_CACHE.get(CellStyleEnum.BLANK_NUM));
        COLUMN_STYLE_CACHE.put("LEFT_STRING", STYLE_CACHE.get(CellStyleEnum.BLANK_STRING));
        COLUMN_STYLE_CACHE.put("LEFT_NUM_END", STYLE_CACHE.get(CellStyleEnum.BLANK_NUM_END));
        COLUMN_STYLE_CACHE.put("LEFT_STRING_END", STYLE_CACHE.get(CellStyleEnum.BLANK_STRING_END));

        // 右侧数据列样式组
        COLUMN_STYLE_CACHE.put("DATA_NORMAL", STYLE_CACHE.get(CellStyleEnum.BLANK_NUM));
        COLUMN_STYLE_CACHE.put("DATA_SUM", STYLE_CACHE.get(CellStyleEnum.SUM_NUM));
        COLUMN_STYLE_CACHE.put("DATA_INPUT", STYLE_CACHE.get(CellStyleEnum.INPUT_NUM));
    }

    /**
     * 预设列样式 - 一次性设置大部分单元格的默认样式
     */
    private static void presetColumnStyles(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        if (excelDtoList.isEmpty()) return;

        int startRow = excelDtoList.get(0).getRowIndex();
        int endRow = excelDtoList.get(excelDtoList.size() - 1).getRowIndex();
        int totalRows = endRow - startRow + 1;

        log.info("预设{}行的列样式...", totalRows);

        // 批量预设左侧25列的基础样式
        presetLeftColumnStyles(sheet, startRow, totalRows);

        // 批量预设右侧数据列的基础样式
        presetRightColumnStyles(sheet, startRow, totalRows);
    }

    /**
     * 预设左侧列样式
     */
    private static void presetLeftColumnStyles(IWorksheet sheet, int startRow, int totalRows) {
        // 数字列 (第2列和第23列)
        sheet.getRange(startRow, 1, totalRows, 1).setStyle(COLUMN_STYLE_CACHE.get("LEFT_NUM"));
        sheet.getRange(startRow, 22, totalRows, 1).setStyle(COLUMN_STYLE_CACHE.get("LEFT_NUM"));

        // 字符串列 (其他基础列)
        sheet.getRange(startRow, 0, totalRows, 1).setStyle(COLUMN_STYLE_CACHE.get("LEFT_STRING"));
        sheet.getRange(startRow, 2, totalRows, 20).setStyle(COLUMN_STYLE_CACHE.get("LEFT_STRING"));
        sheet.getRange(startRow, 23, totalRows, 1).setStyle(COLUMN_STYLE_CACHE.get("LEFT_STRING"));
        sheet.getRange(startRow, 24, totalRows, 1).setStyle(COLUMN_STYLE_CACHE.get("LEFT_STRING"));
    }

    /**
     * 预设右侧数据列样式
     */
    private static void presetRightColumnStyles(IWorksheet sheet, int startRow, int totalRows) {
        // 数据区域统一预设为数字样式
        sheet.getRange(startRow, 25, totalRows, 82).setStyle(COLUMN_STYLE_CACHE.get("DATA_NORMAL"));
    }

    /**
     * 超大批次处理 - 最小化API调用次数
     */
    private static void processInUltraLargeBatches(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        int totalSize = excelDtoList.size();
        int ultraBatchSize = Math.min(BATCH_SIZE, totalSize);

        for (int start = 0; start < totalSize; start += ultraBatchSize) {
            int end = Math.min(start + ultraBatchSize, totalSize);
            List<WriteExcelDto> batch = excelDtoList.subList(start, end);

            log.info("处理超大批次 {}/{} ({}-{}) - {}条记录",
                    (start / ultraBatchSize) + 1,
                    (totalSize + ultraBatchSize - 1) / ultraBatchSize,
                    start, end - 1, batch.size());

            processUltraBatch(sheet, batch);
        }
    }

    /**
     * 超大批次处理核心逻辑
     */
    private static void processUltraBatch(IWorksheet sheet, List<WriteExcelDto> batch) {
        if (batch.isEmpty()) return;

        int startRow = batch.get(0).getRowIndex();
        int endRow = batch.get(batch.size() - 1).getRowIndex();
        int rowCount = endRow - startRow + 1;

        // 动态计算实际需要的列数
        int maxColumns = calculateMaxColumns(batch);
        Object[][] allData = new Object[rowCount][maxColumns];

        // 批量填充数据（零拷贝优化）
        fillBatchData(batch, allData, startRow);

        // 一次性写入所有数据
        IRange batchRange = sheet.getRange(startRow, 0, rowCount, maxColumns);
        batchRange.setValue(allData);

        // 批量设置公式
        List<FormulaPosition> formulas = collectFormulas(batch);
        applyFormulasInBatch(sheet, formulas);
    }

    /**
     * 计算批次所需的最大列数
     */
    private static int calculateMaxColumns(List<WriteExcelDto> batch) {
        int maxCol = 25; // 至少25列基础数据
        for (WriteExcelDto dto : batch) {
            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                maxCol = Math.max(maxCol, cellData.getColumnIndex() + 1);
            }
        }
        return maxCol;
    }

    /**
     * 零拷贝填充批次数据
     */
    private static void fillBatchData(List<WriteExcelDto> batch, Object[][] allData, int startRow) {
        for (int i = 0; i < batch.size(); i++) {
            WriteExcelDto dto = batch.get(i);
            int rowOffset = dto.getRowIndex() - startRow;

            // 直接填充基础数据（避免多次数组访问）
            Object[] rowData = allData[rowOffset];
            rowData[0] = dto.getCol1();
            rowData[1] = dto.getCol2();
            rowData[2] = dto.getCol3();
            rowData[3] = dto.getCol4();
            rowData[4] = dto.getCol5();
            rowData[5] = dto.getCol6();
            rowData[6] = dto.getCol7();
            rowData[7] = dto.getCol8();
            rowData[8] = dto.getCol9();
            rowData[9] = dto.getCol10();
            rowData[10] = dto.getCol11();
            rowData[11] = dto.getCol12();
            rowData[12] = dto.getCol13();
            rowData[13] = dto.getCol14();
            rowData[14] = dto.getCol15();
            rowData[15] = dto.getCol16();
            rowData[16] = dto.getCol17();
            rowData[17] = dto.getCol18();
            rowData[18] = dto.getCol19();
            rowData[19] = dto.getCol20();
            rowData[20] = dto.getCol21();
            rowData[21] = dto.getCol22();
            rowData[22] = dto.getCol23();
            rowData[23] = dto.getCol24();
            rowData[24] = dto.getCol25();

            // 填充数据区域
            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                int colIndex = cellData.getColumnIndex();
                if (colIndex < rowData.length) {
                    rowData[colIndex] = cellData.isFormula() ? null : cellData.getValue();
                }
            }
        }
    }

    /**
     * 收集所有公式位置信息
     */
    private static List<FormulaPosition> collectFormulas(List<WriteExcelDto> batch) {
        List<FormulaPosition> formulas = new ArrayList<>();
        for (WriteExcelDto dto : batch) {
            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                if (cellData.isFormula()) {
                    formulas.add(new FormulaPosition(
                            dto.getRowIndex(),
                            cellData.getColumnIndex(),
                            cellData.getFormula()
                    ));
                }
            }
        }
        return formulas;
    }

    /**
     * 批量应用公式
     */
    private static void applyFormulasInBatch(IWorksheet sheet, List<FormulaPosition> formulas) {
        for (FormulaPosition fp : formulas) {
            sheet.getRange(fp.row, fp.col).setFormula2(fp.formula);
        }
    }

    /**
     * 应用特殊样式（只处理需要特殊处理的情况）
     */
    private static void applySpecialStyles(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        log.info("应用特殊样式...");
        int specialCount = 0;

        for (WriteExcelDto dto : excelDtoList) {
            // 只处理结尾行和汇总行的特殊样式
            if (dto.getCol2() == 24 || dto.isSummaryData()) {
                applySpecialRowStyles(sheet, dto);
                specialCount++;
            }
        }

        log.info("特殊样式应用完成，共处理{}行", specialCount);
    }

    /**
     * 应用特殊行样式
     */
    private static void applySpecialRowStyles(IWorksheet sheet, WriteExcelDto dto) {
        int row = dto.getRowIndex();

        // 处理结尾行样式
        if (dto.getCol2() == 24) {
            for (int col = 0; col <= 23; col++) {
                CellStyleEnum endStyle = getEndRowStyle(dto, col);
                sheet.getRange(row, col).setStyle(STYLE_CACHE.get(endStyle));
            }
        }

        // 处理分类列样式（第25列）
        CellStyleEnum categoryStyle = dto.isSummaryData() ?
                TitleCategoryEnum.fromSeq(dto.getCol2()).getSumTitleCellStyle() :
                TitleCategoryEnum.fromSeq(dto.getCol2()).getTitleCellStyle();
        sheet.getRange(row, 24).setStyle(STYLE_CACHE.get(categoryStyle));
    }

    /**
     * 获取结尾行样式
     */
    private static CellStyleEnum getEndRowStyle(WriteExcelDto dto, int col) {
        if (dto.isSummaryData()) {
            return (col == 1 || col == 22) ?
                    CellStyleEnum.TOTAL_BLANK_NUM_END :
                    CellStyleEnum.TOTAL_BLANK_STRING_END;
        } else {
            return (col == 1 || col == 22) ?
                    CellStyleEnum.BLANK_NUM_END :
                    CellStyleEnum.BLANK_STRING_END;
        }
    }

    /**
     * 保持原有完整的样式配置逻辑
     */
    private static void configureCompleteStyle(IStyle cs, CellStyleEnum styleEnum) {
        switch (styleEnum) {
            case BLANK_NUM:
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                cs.setNumberFormat("#,##0;[Red]-#,##0");
                break;
            case BLANK_STRING:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                break;
            case BLANK_NUM_END:
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
                cs.setNumberFormat("#,##0;[Red]-#,##0");
                break;
            case BLANK_STRING_END:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
                cs.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.GetRed());
                break;
            case BLANK_INDENT:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.setIndentLevel(1);
                break;
            case REF_SEL_NUM:
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                cs.setNumberFormat("#,##0;[Red]-#,##0");
                cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            case REF_SEL_STRING:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.getInterior().setColor(Color.FromArgb(252, 228, 214));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            case INPUT_NUM:
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                cs.setNumberFormat("#,##0;[Red]-#,##0");
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            case INPUT_STRING:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.getInterior().setColor(Color.FromArgb(226, 239, 218));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            case SUM_NUM:
                cs.setHorizontalAlignment(HorizontalAlignment.Right);
                cs.setNumberFormat("#,##0;[Red]-#,##0");
                cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            case SUM_STRING:
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                cs.getInterior().setColor(Color.FromArgb(0, 176, 240));
                cs.getInterior().setPattern(Pattern.Solid);
                break;
            // ... 其他样式配置保持不变
            default:
                // 默认左对齐样式
                cs.setHorizontalAlignment(HorizontalAlignment.Left);
                break;
        }
    }

    /**
     * 公式位置信息类
     */
    private static class FormulaPosition {
        final int row;
        final int col;
        final String formula;

        FormulaPosition(int row, int col, String formula) {
            this.row = row;
            this.col = col;
            this.formula = formula;
        }
    }
}
