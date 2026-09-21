package com.example.demo2.controller;

import com.example.demo2.GcexcelCopyStyle;
import com.example.demo2.entity.WriteExcelDto;
import com.grapecity.documents.excel.*;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.net.http.HttpResponse;
import java.util.*;

@Slf4j
@RestController
@RequestMapping("/api/excel")
public class ExcelDownCpntroller {

    public static final String PSI_TEMPLATE_FILE = "templates/excel/sheet.xlsx";

    // 缓存样式对象，避免重复创建
    private static final Map<String, IStyle> STYLE_CACHE = new HashMap<>();

    /**
     * 下载Excel文件接口
     * @return ResponseEntity包含Excel文件
     */
    @GetMapping("/download")
    public void downloadExcel(HttpServletResponse response) {
        long startTime = System.currentTimeMillis();

        try {
            // 设置响应头
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment; filename=\"generated_excel.xlsx\"");
            response.setCharacterEncoding("UTF-8");

            // 生成Excel并直接写入响应输出流
            generateExcel(response.getOutputStream());

            long endTime = System.currentTimeMillis();
            log.info("Excel下载完成，耗时: {} ms", endTime - startTime);
        } catch (Exception e) {
            log.error("生成Excel文件失败", e);
            try {
                response.setStatus(HttpStatus.INTERNAL_SERVER_ERROR.value());
                response.getWriter().write("Excel生成失败: " + e.getMessage());
            } catch (IOException ioException) {
                log.error("返回错误信息失败", ioException);
            }
        }
    }

    /**
     * 生成Excel文件的核心方法
     * @return Excel文件的字节数组
     */
    private void generateExcel(OutputStream outputStream) throws IOException {
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
                if ( i%2 == 0) {
                    writeExcelDto.setCol24(null);
                } else {
                    writeExcelDto.setCol24(70L);
                }

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
//            initializeStyleCache(workbook);

            log.info("开始生成excel，共{}行数据", excelDtoList.size());

            // 批量处理 - 更大的批次
            processInLargeBatches(sheet, excelDtoList);

            // 在所有数据处理完成后应用样式
            if (!excelDtoList.isEmpty()) {
                int firstRow = excelDtoList.get(0).getRowIndex();
                int lastRow = excelDtoList.get(excelDtoList.size() - 1).getRowIndex();
                int totalRowCount = lastRow - firstRow + 1;
                log.info("开始统一应用样式，处理 {} 行数据", totalRowCount);
                applyBatchStyles(sheet, excelDtoList, firstRow, totalRowCount);
            }

            log.info("excel生成完毕");

            // 直接写入响应输出流
            workbook.save(outputStream, SaveFileFormat.Xlsx);

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

            // 确保输出流被刷新和关闭
            try {
                outputStream.flush();
                outputStream.close();
            } catch (IOException e) {
                log.warn("关闭输出流失败", e);
            }
        }
    }

        /**
         * 初始化样式缓存
         */
        private  void initializeStyleCache(Workbook workbook) {
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
        private  void processInLargeBatches(IWorksheet sheet, List<WriteExcelDto> excelDtoList) {
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
        private  void processBatchOptimized(IWorksheet sheet, List<WriteExcelDto> batch) {
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

        }

        /**
         * 批量应用样式
         */
        private  void applyBatchStyles(IWorksheet sheet, List<WriteExcelDto> batch,
        int startRow, int rowCount) {
            var pasteOption = new PasteOption();
            pasteOption.setPasteType(EnumSet.of(PasteType.Formats));

            // 计算有多少组数据（每组24行）
            int groupCount = (int) Math.ceil((double) rowCount / 24);

            log.info("需要复制样式 {} 组，每组24行", groupCount);

            // 源样式区域：A4:DC27 (24行)
            IRange sourceRange = sheet.getRange("A4:DC27");

            // 为每组数据复制样式
            for (int group = 0; group < groupCount; group++) {
                int targetStartRow = startRow + (group * 24);
                int rowsInThisGroup = Math.min(24, rowCount - (group * 24));

                if (rowsInThisGroup > 0) {
                    // 目标区域：从targetStartRow开始，复制rowsInThisGroup行

                    String targetRangeAddress = String.format("A%d:DC%d",
                            targetStartRow+1, targetStartRow + rowsInThisGroup);
                    IRange targetRange = sheet.getRange(targetRangeAddress);
                    sourceRange.copy(targetRange, pasteOption);

                    log.info("复制样式到组 {}: {} ({}行)", group + 1, targetRangeAddress, rowsInThisGroup);
                }
            }
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
}
