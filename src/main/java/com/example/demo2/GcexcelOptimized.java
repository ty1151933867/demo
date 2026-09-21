package com.example.demo2;

import com.example.demo2.entity.CellStyleEnum;
import com.example.demo2.entity.TitleCategoryEnum;
import com.example.demo2.entity.WriteExcelDto;
import com.grapecity.documents.excel.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Slf4j
public class GcexcelOptimized {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    // 缓存样式对象，避免重复创建
    private static final Map<String, IStyle> STYLE_CACHE = new HashMap<>();

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        int rowIndexCounter = 2;
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

            // 极致性能优化设置
            workbook.setEnableCalculation(false);
            workbook.setDeferUpdateDirtyState(true);
            workbook.setReferenceStyle(ReferenceStyle.A1); // 使用A1引用样式

            workbook.open(inputStream);

            IWorksheet sheet = workbook.getWorksheets().get(0);

            // 预热样式缓存
            initializeStyleCache(workbook);

            log.info("开始生成excel，共{}行数据", excelDtoList.size());

            // 批量处理 - 更大的批次
            processInLargeBatches(sheet, excelDtoList);

            // 清除实际使用以外的区域
            clearUnusedAreas(sheet, excelDtoList);

            log.info("excel生成完毕");
            try (FileOutputStream fos = new FileOutputStream("test_Gcexcel_optimized.xlsx")) {
                workbook.save(fos);
                log.info("Excel文件已保存到: test_Gcexcel_optimized.xlsx");
            } catch (IOException e) {
                log.error("保存文件时发生错误", e);
                throw new RuntimeException("保存Excel文件失败", e);
            }


        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            long endTime = System.currentTimeMillis();
            long totalTime = endTime - startTime;
            log.info("Excel生成完成，总耗时: {} ms", totalTime);
            System.out.println("Excel生成完成，总耗时: " + totalTime + " ms (" + (totalTime / 1000.0) + " 秒)");

//            if (workbook != null) {
//                try {
//                    workbook.close();
//                } catch (Exception e) {
//                    log.warn("关闭工作簿失败", e);
//                }
//            }
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
     * 初始化样式缓存
     */
    private static void initializeStyleCache(Workbook workbook) {
        log.info("初始化样式缓存...");

        // 基础样式
        IStyle baseStyle = workbook.getStyles().add("BaseStyle");
        baseStyle.getFont().setName("Meiryo UI");
        baseStyle.getFont().setSize(9);
        baseStyle.setVerticalAlignment(VerticalAlignment.Center);

        // 数字样式
        IStyle numStyle = workbook.getStyles().add("NumStyle");
        numStyle.getFont().setName("Meiryo UI");
        numStyle.getFont().setSize(9);
        numStyle.setVerticalAlignment(VerticalAlignment.Center);
        numStyle.setHorizontalAlignment(HorizontalAlignment.Right);
        numStyle.setNumberFormat("#,##0;[Red]-#,##0");

        // 字符串样式
        IStyle stringStyle = workbook.getStyles().add("StringStyle");
        stringStyle.getFont().setName("Meiryo UI");
        stringStyle.getFont().setSize(9);
        stringStyle.setVerticalAlignment(VerticalAlignment.Center);
        stringStyle.setHorizontalAlignment(HorizontalAlignment.Left);

        STYLE_CACHE.put("BASE", baseStyle);
        STYLE_CACHE.put("NUM", numStyle);
        STYLE_CACHE.put("STRING", stringStyle);

        log.info("样式缓存初始化完成");
    }

    /**
     * 大批次处理
     */
    private static void processInLargeBatches(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        int batchSize = 2000; // 增大批处理大小
        int totalSize = excelDtoList.size();

        for (int start = 0; start < totalSize; start += batchSize) {
            int end = Math.min(start + batchSize, totalSize);
            List<WriteExcelDto> batch = excelDtoList.subList(start, end);

            log.info("处理大批次 {}/{} ({}-{})",
                    (start / batchSize) + 1,
                    (totalSize + batchSize - 1) / batchSize,
                    start, end - 1);

            // 批量处理整个批次的数据
            processBatchOptimized(sheet, batch);
        }
    }

    /**
     * 优化的批次处理 - 最小化API调用
     */
    private static void processBatchOptimized(IWorksheet sheet, List<WriteExcelDto> batch) {
        if (batch.isEmpty()) return;

        int startRow = batch.get(0).getRowIndex();
        int endRow = batch.get(batch.size() - 1).getRowIndex();
        int rowCount = endRow - startRow + 1;

        // 获取第一个元素的dataMap大小来动态确定列数
        int dataMapSize = batch.get(0).getDataMap().size();
        int totalColumns = 25 + dataMapSize; // 25基础列 + 动态dataMap大小

        // 批量准备所有数据（基于实际的dataMap大小）
        Object[][] allData = new Object[rowCount][totalColumns];

        for (int i = 0; i < batch.size(); i++) {
            WriteExcelDto dto = batch.get(i);
            int rowOffset = dto.getRowIndex() - startRow;

            // 填充基础数据
            allData[rowOffset][0] = dto.getCol1();
            allData[rowOffset][1] = dto.getCol2();
            allData[rowOffset][2] = dto.getCol3();
            allData[rowOffset][3] = dto.getCol4();
            allData[rowOffset][4] = dto.getCol5();
            allData[rowOffset][5] = dto.getCol6();
            allData[rowOffset][6] = dto.getCol7();
            allData[rowOffset][7] = dto.getCol8();
            allData[rowOffset][8] = dto.getCol9();
            allData[rowOffset][9] = dto.getCol10();
            allData[rowOffset][10] = dto.getCol11();
            allData[rowOffset][11] = dto.getCol12();
            allData[rowOffset][12] = dto.getCol13();
            allData[rowOffset][13] = dto.getCol14();
            allData[rowOffset][14] = dto.getCol15();
            allData[rowOffset][15] = dto.getCol16();
            allData[rowOffset][16] = dto.getCol17();
            allData[rowOffset][17] = dto.getCol18();
            allData[rowOffset][18] = dto.getCol19();
            allData[rowOffset][19] = dto.getCol20();
            allData[rowOffset][20] = dto.getCol21();
            allData[rowOffset][21] = dto.getCol22();
            allData[rowOffset][22] = dto.getCol23();
            allData[rowOffset][23] = dto.getCol24();
            allData[rowOffset][24] = dto.getCol25();

            // 填充数据区域
            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                int colIndex = cellData.getColumnIndex();
                if (colIndex >= 25 && colIndex < totalColumns) {
                    allData[rowOffset][colIndex] = cellData.isFormula() ? null : cellData.getValue();
                }
            }
        }

        // 一次性写入所有数据
        IRange batchRange = sheet.getRange(startRow, 0, rowCount, totalColumns);
        batchRange.setValue(allData);

        // 批量设置公式（如果有）
        List<FormulaInfo> formulas = new ArrayList<>();
        for (int i = 0; i < batch.size(); i++) {
            WriteExcelDto dto = batch.get(i);
            int rowOffset = dto.getRowIndex() - startRow;

            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                if (cellData.isFormula()) {
                    formulas.add(new FormulaInfo(
                            dto.getRowIndex(),
                            cellData.getColumnIndex(),
                            cellData.getFormula()
                    ));
                }
            }
        }

        // 批量设置公式
        for (FormulaInfo formula : formulas) {
            sheet.getRange(formula.row, formula.col).setFormula2(formula.formula);
        }

        // 批量设置样式（简化版）
//        applyBatchStyles(sheet, batch, startRow, rowCount);
    }

    /**
     * 批量应用样式
     */
    private static void applyBatchStyles(IWorksheet sheet, List<WriteExcelDto> batch,
                                         int startRow, int rowCount) {
        // 基础列样式
        sheet.getRange(startRow, 1, rowCount, 1).setStyle(STYLE_CACHE.get("NUM")); // 数字列
        sheet.getRange(startRow, 22, rowCount, 1).setStyle(STYLE_CACHE.get("NUM")); // 数字列

        // 其他列使用字符串样式
        sheet.getRange(startRow, 0, rowCount, 1).setStyle(STYLE_CACHE.get("STRING"));
        sheet.getRange(startRow, 2, rowCount, 20).setStyle(STYLE_CACHE.get("STRING"));
        sheet.getRange(startRow, 23, rowCount, 1).setStyle(STYLE_CACHE.get("STRING"));
        sheet.getRange(startRow, 24, rowCount, 1).setStyle(STYLE_CACHE.get("STRING"));

        // 数据区域样式
        sheet.getRange(startRow, 25, rowCount, 80).setStyle(STYLE_CACHE.get("NUM"));
    }

    // 公式信息辅助类
    private static class FormulaInfo {
        final int row;
        final int col;
        final String formula;

        FormulaInfo(int row, int col, String formula) {
            this.row = row;
            this.col = col;
            this.formula = formula;
        }
    }

    /**
     * 清除实际使用以外的区域
     * 删除未使用的行和列，减小文件大小
     */
    private static void clearUnusedAreas(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
        if (excelDtoList.isEmpty()) return;

        log.info("开始清理未使用的区域...");

        // 获取实际使用的最大行列数
        int maxUsedRow = excelDtoList.get(excelDtoList.size() - 1).getRowIndex();
        int maxUsedCol = 25; // 基础列数

        // 计算数据区域的最大列数
        for (WriteExcelDto dto : excelDtoList) {
            for (WriteExcelDto.CellData cellData : dto.getDataMap().values()) {
                maxUsedCol = Math.max(maxUsedCol, cellData.getColumnIndex());
            }
        }

        // 获取工作表的实际行列数
        int totalRows = sheet.getUsedRange().getRowCount();
        int totalCols = sheet.getUsedRange().getColumnCount();

        log.info("实际使用范围: 行[1-{}], 列[1-{}]", maxUsedRow, maxUsedCol);
        log.info("工作表总范围: 行[1-{}], 列[1-{}]", totalRows, totalCols);

        // 清除未使用的行（从最大使用行+1开始到最后一行）
        if (maxUsedRow < totalRows - 1) {
            IRange unusedRows = sheet.getRange(maxUsedRow + 1, 0, totalRows - maxUsedRow, totalCols);
            unusedRows.clear();
            log.info("已清除 {} 行未使用的行", totalRows - maxUsedRow - 1);
        }

        // 清除未使用的列（从最大使用列+1开始到最后一列）
        if (maxUsedCol < totalCols - 1) {
            IRange unusedCols = sheet.getRange(0, maxUsedCol + 1, maxUsedRow, totalCols - maxUsedCol);
            unusedCols.clear();
            log.info("已清除 {} 列未使用的列", totalCols - maxUsedCol - 1);
        }
        log.info("未使用区域清理完成");
    }

}
