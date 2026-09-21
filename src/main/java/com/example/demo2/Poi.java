package com.example.demo2;

import com.example.demo2.entity.CellStyleEnum;
import com.example.demo2.entity.TitleCategoryEnum;
import com.example.demo2.entity.WriteExcelDto;
import com.example.demo2.utils.Utils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Slf4j
public class Poi {

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
        XSSFWorkbook workbookTemp = null;
        SXSSFWorkbook sxssfWorkbook = null;

        try {
            inputStream = new ClassPathResource(PSI_TEMPLATE_FILE).getInputStream();
            workbookTemp = new XSSFWorkbook(inputStream);
            sxssfWorkbook = new SXSSFWorkbook(workbookTemp, SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
            sxssfWorkbook.setForceFormulaRecalculation(true);
            SXSSFSheet sheet = sxssfWorkbook.getSheetAt(0);

            // style生成
            Map<CellStyleEnum, CellStyle> cellStyleMap = getDataCellStyleMap(sxssfWorkbook);

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

                    Cell cell =Utils.getCell(sheet, cellData.getRowIndex(), cellData.getColumnIndex());
                    // 有公式
                    if (cellData.isFormula()){
                        // 有公式
                        cell.setCellFormula(cellData.getFormula());
                    } else {
                        Utils.setCellValue(cell, cellData.getValue());
                    }

                    // style设置样式
                    if (writeExcelDto.isSummaryData()) {
                        cell.setCellStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getSumTitleCellStyle()));
                    } else {
                        cell.setCellStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(writeExcelDto.getCol2()).getTitleCellStyle()));
                    }

                }
                log.info("第{}行右侧数据写入结束", i);
            }
            log.info("excel生成完毕");
            saveExcelFile(sxssfWorkbook, "test_poi.xlsx");


        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            long endTime = System.currentTimeMillis(); // 结束时间
            long totalTime = endTime - startTime; // 总耗时（毫秒）
            log.info("Excel生成完成，总耗时: {} ms", totalTime);
            System.out.println("Excel生成完成，总耗时: " + totalTime + " ms (" + (totalTime / 1000.0) + " 秒)");
        }
    }

    private static void saveExcelFile(SXSSFWorkbook workbook, String fileName) throws IOException {
        java.io.FileOutputStream fileOut = new java.io.FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();
        log.info("Excel文件已保存到: {}", new java.io.File(fileName).getAbsolutePath());
    }

        private static void setBasicInfo(SXSSFSheet sheet, WriteExcelDto basicData,
                                         Map<CellStyleEnum, CellStyle> cellStyleMap, boolean isEnd) {

        Cell alertFlagCell = Utils.getCell(sheet, basicData.getRowIndex(),0);
        Utils.setCellValue(alertFlagCell, basicData.getCol1());
        if (basicData.isSummaryData()) {
            alertFlagCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            alertFlagCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }

        Cell idCell = Utils.getCell(sheet, basicData.getRowIndex(),1);
        Utils.setCellValue(idCell, basicData.getCol2());
        if (basicData.isSummaryData()) {
            idCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM));
        } else {
            idCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_NUM));
        }

        Cell sortKeyCell = Utils.getCell(sheet, basicData.getRowIndex(),2);
        Utils.setCellValue(sortKeyCell, basicData.getCol3());
        if (basicData.isSummaryData()) {
            sortKeyCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            sortKeyCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell divCell = Utils.getCell(sheet, basicData.getRowIndex(), 3);
        Utils.setCellValue(divCell, basicData.getCol4());
        if (basicData.isSummaryData()) {
            divCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            divCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell subCell = Utils.getCell(sheet, basicData.getRowIndex(), 4);
        Utils.setCellValue(subCell, basicData.getCol5());
        if (basicData.isSummaryData()) {
            subCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            subCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell consoleCell = Utils.getCell(sheet, basicData.getRowIndex(),5);
        Utils.setCellValue(consoleCell, basicData.getCol6());
        if (basicData.isSummaryData()) {
            consoleCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            consoleCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell itemConsolCell = Utils.getCell(sheet, basicData.getRowIndex(),6);
        Utils.setCellValue(itemConsolCell, basicData.getCol7());
        if (basicData.isSummaryData()) {
            itemConsolCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemConsolCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell itemCodeCell = Utils.getCell(sheet, basicData.getRowIndex(), 7);
        Utils.setCellValue(itemCodeCell, basicData.getCol8());
        if (basicData.isSummaryData()) {
            itemCodeCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemCodeCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell itemNameSymbolCell = Utils.getCell(sheet, basicData.getRowIndex(),8);
        Utils.setCellValue(itemNameSymbolCell, basicData.getCol9());
        if (basicData.isSummaryData()) {
            itemNameSymbolCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemNameSymbolCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell itemNameCell = Utils.getCell(sheet, basicData.getRowIndex(),9);
        Utils.setCellValue(itemNameCell, basicData.getCol10());
        if (basicData.isSummaryData()) {
            itemNameCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            itemNameCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell proPlantCell = Utils.getCell(sheet, basicData.getRowIndex(),10);
        Utils.setCellValue(proPlantCell, basicData.getCol11());
        if (basicData.isSummaryData()) {
            proPlantCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            proPlantCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell dosageCell = Utils.getCell(sheet, basicData.getRowIndex(),11);
        Utils.setCellValue(dosageCell, basicData.getCol12());
        if (basicData.isSummaryData()) {
            dosageCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            dosageCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell minInvCell = Utils.getCell(sheet, basicData.getRowIndex(),12);
        Utils.setCellValue(minInvCell, basicData.getCol13());
        if (basicData.isSummaryData()) {
            minInvCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            minInvCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell maxInvCell = Utils.getCell(sheet, basicData.getRowIndex(),13);
        Utils.setCellValue(maxInvCell, basicData.getCol14());
        if (basicData.isSummaryData()) {
            maxInvCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            maxInvCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell minInvQtyCell = Utils.getCell(sheet, basicData.getRowIndex(),14);
        Utils.setCellValue(minInvQtyCell, basicData.getCol15());
        if (basicData.isSummaryData()) {
            minInvQtyCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            minInvQtyCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell weeklyPrdCell = Utils.getCell(sheet, basicData.getRowIndex(),15 );
        Utils.setCellValue(weeklyPrdCell, basicData.getCol16());
        if (basicData.isSummaryData()) {
            weeklyPrdCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            weeklyPrdCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell moqCell = Utils.getCell(sheet, basicData.getRowIndex(),16);
        Utils.setCellValue(moqCell, basicData.getCol17());
        if (basicData.isSummaryData()) {
            moqCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            moqCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell lotCell = Utils.getCell(sheet, basicData.getRowIndex(),17);
        Utils.setCellValue(lotCell, basicData.getCol18());
        if (basicData.isSummaryData()) {
            lotCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            lotCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell orderLt = Utils.getCell(sheet, basicData.getRowIndex(), 18);
        Utils.setCellValue(orderLt, basicData.getCol19());
        if (basicData.isSummaryData()) {
            orderLt.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            orderLt.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell costCell = Utils.getCell(sheet, basicData.getRowIndex(),19);
        Utils.setCellValue(costCell, basicData.getCol20());
        if (basicData.isSummaryData()) {
            costCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            costCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell qtyPerPkgCell = Utils.getCell(sheet, basicData.getRowIndex(),20);
        Utils.setCellValue(qtyPerPkgCell, basicData.getCol21());
        if (basicData.isSummaryData()) {
            qtyPerPkgCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            qtyPerPkgCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell isAggCell = Utils.getCell(sheet, basicData.getRowIndex(),21);
        Utils.setCellValue(isAggCell, basicData.getCol22());
        if (basicData.isSummaryData()) {
            isAggCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            isAggCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell sortCell = Utils.getCell(sheet, basicData.getRowIndex(), 22);
        Utils.setCellValue(sortCell, basicData.getCol23());
        if (basicData.isSummaryData()) {
            sortCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_NUM));
        } else {
            sortCell.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_NUM_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_NUM));
        }


        Cell psiType = Utils.getCell(sheet, basicData.getRowIndex(), 23);
        Utils.setCellValue(psiType, basicData.getCol24());
        if (basicData.isSummaryData()) {
            psiType.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.TOTAL_BLANK_STRING));
        } else {
            psiType.setCellStyle(isEnd ? cellStyleMap.get(CellStyleEnum.BLANK_STRING_END)
                    : cellStyleMap.get(CellStyleEnum.BLANK_STRING));
        }


        Cell caterogyNameCell = Utils.getCell(sheet, basicData.getRowIndex(), 24);
        Utils.setCellValue(caterogyNameCell, basicData.getCol25());
        if (basicData.isSummaryData()) {
            caterogyNameCell.setCellStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getSumTitleCellStyle()));
        } else {
            caterogyNameCell.setCellStyle(cellStyleMap.get(TitleCategoryEnum.fromSeq(basicData.getCol2()).getTitleCellStyle()));
        }
    }


    private static Map<CellStyleEnum, CellStyle> getDataCellStyleMap(org.apache.poi.ss.usermodel.Workbook workbook) {
        Map<CellStyleEnum, CellStyle> dataCellStyleMap = new HashMap<>();

        DataFormat df = workbook.createDataFormat();
        // フォントとフォントサイズ
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("Meiryo UI");

        for (CellStyleEnum styleEnum : CellStyleEnum.values()) {
            CellStyle cs = workbook.createCellStyle();

            cs.setFont(font);
            cs.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

            if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_NUM_END.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());

                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_STRING_END.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());
            } else if (styleEnum.getValue() == CellStyleEnum.BLANK_INDENT.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
                cs.setIndention((short) 1);
            } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 252, (byte) 228, (byte) 214 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.REF_SEL_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 252, (byte) 228, (byte) 214 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 226, (byte) 239, (byte) 218 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 226, (byte) 239, (byte) 218 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.SUM_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 0, (byte) 176, (byte) 240 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.SUM_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 0, (byte) 176, (byte) 240 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_NUM_END.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 226, (byte) 239, (byte) 218 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.INPUT_STRING_END.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 226, (byte) 239, (byte) 218 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_NUM_END.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());

                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_STRING_END.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // ボーダーラインを設定
                cs.setBorderBottom(BorderStyle.THIN);
                cs.setBottomBorderColor(IndexedColors.RED.getIndex());

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_BLANK_INDENT.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
                cs.setIndention((short) 1);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_NUM.getValue()) {
                // 右詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                // データ フォーマット
                cs.setDataFormat(df.getFormat("#,##0;[Red]-#,##0"));

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_SUM_STRING.getValue()) {
                // 左詰
                cs.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if (styleEnum.getValue() == CellStyleEnum.TOTAL_REF_SEL_STRING.getValue()) {
                // 左詰
                cs.setAlignment(HorizontalAlignment.LEFT);

                // 背景色を設定
                cs.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 217, (byte) 217, (byte) 217 }, null));
                cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            dataCellStyleMap.put(styleEnum, cs);
        }

        return dataCellStyleMap;
    }

}
