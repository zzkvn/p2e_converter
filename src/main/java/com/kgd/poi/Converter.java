package com.kgd.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xslf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Converter {
    private List<Record> records = new ArrayList<>();
    private static final String HEADER_PROJECT_CODE = "项目代号";
    private static final String HEADER_NBA = "NBA";
    private static final String HEADER_PRODUCT_PRICE = "产品价格";
    private static final String OUTPUT_FILE_NAME = "2023.xx.xx上定点会-汇总信息";

    public String convert(File[] pptFiles, File excelFileDir) throws IOException {
        try {
            parsePptFiles(pptFiles);
            return writeExcelFile(excelFileDir);
        } catch (IOException e) {
            throw e;
        }
    }

    private void parsePptFiles(File[] pptFiles) throws IOException {
        for (File pptFile : pptFiles) {
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile));
            XSLFTable table;
            for (XSLFSlide slide : ppt.getSlides()) {
                System.out.println(slide.getSlideName());
                XSLFTable table1 = null;
                XSLFTable table2 = null;
                for (XSLFShape shape : slide) {
                    if (shape instanceof XSLFTable) {
                        table = (XSLFTable) shape;
                        for (XSLFTableRow row : table.getRows()) {
                            String firstCell = row.getCells().get(0).getText();
                            if (firstCell.equals(HEADER_PROJECT_CODE)) {
                                table1 = table;
                                break;
                            } else if (firstCell.contains(HEADER_NBA) && firstCell.contains(HEADER_PRODUCT_PRICE)) {
                                table2 = table;
                                break;
                            }
                        }
                    }
                }
                if (table1 != null && table2 != null) {
                    List<Record> recordsInSlide = processTable1(table1);
                    processTable2(table2, recordsInSlide);
                    records.addAll(recordsInSlide);
                } else {
                    System.out.println("Table1 or Table2 not found");
                }
            }
        }
        System.out.println(records.size());
    }

    private List<Record> processTable1(XSLFTable table1) {
        List<Record> records = new ArrayList<>();
        boolean start = false;
        for (XSLFTableRow row : table1.getRows()) {
            List<XSLFTableCell> cells = row.getCells();
            String firstCell = row.getCells().get(0).getText();
            if (firstCell.equals(HEADER_PROJECT_CODE)) {
                start = true;
                continue;
            } else if (firstCell.isBlank()) {
                continue;
            }
            if (start) {
                Record record = new Record();
                record.setProjectCode(cells.get(0).getText());
                record.setProjectCategory(cells.get(1).getText());
                record.setBuyer(cells.get(2).getText());
                record.setSq(cells.get(3).getText());
                record.setProjectName(cells.get(4).getText());
                record.setAnnualDemand(cells.get(5).getText());
                record.setProductSapCode(cells.get(7).getText());
                record.setProductFactory(cells.get(8).getText());
                record.setProductName(cells.get(9).getText());
                record.setProductMaterial(cells.get(10).getText());
                record.setProductWeight(cells.get(11).getText());
                record.setSsrCompleteDate(cells.get(12).getText());
                record.setRequiredDeliveryDate(cells.get(13).getText());
                records.add(record);
            }
        }
        return records;
    }

    private void processTable2(XSLFTable table2, List<Record> records) {
        boolean start = false;
        Iterator<Record> it = records.listIterator();
        for (XSLFTableRow row : table2.getRows()) {
            List<XSLFTableCell> cells = row.getCells();
            String firstCell = row.getCells().get(0).getText();
            if (firstCell.contains(HEADER_NBA) && firstCell.contains(HEADER_PRODUCT_PRICE)) {
                start = true;
                continue;
            } else if (firstCell.isBlank()) {
                continue;
            }
            if (start) {
                Record record = it.next();
                record.setNbaProductPrice(cells.get(0).getText());
                record.setNbaBudget(cells.get(1).getText());

                record.setSupplier1Info(cells.get(2).getText());
                record.setSupplier1QuotedPrice(cells.get(3).getText());
                record.setSupplier1OtherFee(cells.get(4).getText());
                record.setSupplier1AnnualDecrease(cells.get(5).getText());

                record.setSupplier2Info(cells.get(6).getText());
                record.setSupplier2QuotedPrice(cells.get(7).getText());
                record.setSupplier2OtherFee(cells.get(8).getText());
                record.setSupplier2AnnualDecrease(cells.get(9).getText());

                record.setSupplier3Info(cells.get(10).getText());
                record.setSupplier3QuotedPrice(cells.get(11).getText());
                record.setSupplier3OtherFee(cells.get(12).getText());
                record.setSupplier3AnnualDecrease(cells.get(13).getText());
            }
        }
    }

    private String writeExcelFile(File excelFileDir) throws IOException {
        InputStream is = getClass().getClassLoader().getResourceAsStream("files/Sample.xlsx");
        Workbook workbook = WorkbookFactory.create(is);
        Sheet sheet = workbook.getSheet("上会明细");
        int rowStart = 3;
        for(int i = 0; i < records.size(); i++) {
            Record record = records.get(i);
            Row row = sheet.getRow(rowStart +i);
            row.getCell(3, CREATE_NULL_AS_BLANK).setCellValue(record.getProjectCode());
            row.getCell(4, CREATE_NULL_AS_BLANK).setCellValue(record.getProjectCategory());
            row.getCell(5, CREATE_NULL_AS_BLANK).setCellValue(record.getBuyer());
            row.getCell(6, CREATE_NULL_AS_BLANK).setCellValue(record.getSq());
            row.getCell(8, CREATE_NULL_AS_BLANK).setCellValue(record.getProductName());
            row.getCell(9, CREATE_NULL_AS_BLANK).setCellValue(record.getAnnualDemand());
            row.getCell(10, CREATE_NULL_AS_BLANK).setCellValue(record.getProductSapCode());
            row.getCell(11, CREATE_NULL_AS_BLANK).setCellValue(record.getProductFactory());
            row.getCell(12, CREATE_NULL_AS_BLANK).setCellValue(record.getProductName());
            row.getCell(13, CREATE_NULL_AS_BLANK).setCellValue(record.getProductMaterial());
            row.getCell(14, CREATE_NULL_AS_BLANK).setCellValue(record.getProductWeight());
            row.getCell(15, CREATE_NULL_AS_BLANK).setCellValue(record.getSsrCompleteDate());
            row.getCell(16, CREATE_NULL_AS_BLANK).setCellValue(record.getRequiredDeliveryDate());
            row.getCell(18, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier1Info());
            row.getCell(19, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier1QuotedPrice());
            row.getCell(20, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier1OtherFee());
            row.getCell(21, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier1AnnualDecrease());

            row.getCell(22, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier2Info());
            row.getCell(23, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier2QuotedPrice());
            row.getCell(24, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier2OtherFee());
            row.getCell(25, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier2AnnualDecrease());

            row.getCell(26, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier3Info());
            row.getCell(27, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier3QuotedPrice());
            row.getCell(28, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier3OtherFee());
            row.getCell(29, CREATE_NULL_AS_BLANK).setCellValue(record.getSupplier3AnnualDecrease());
            row.getCell(30, CREATE_NULL_AS_BLANK).setCellValue(record.getProductSapCode());
            row.getCell(32, CREATE_NULL_AS_BLANK).setCellValue(record.getNbaProductPrice());
            row.getCell(34, CREATE_NULL_AS_BLANK).setCellValue(record.getNbaBudget());
        }
        is.close();
        File outputFile = getOutputFileName(excelFileDir);
        FileOutputStream os = new FileOutputStream(outputFile);
        workbook.write(os);
        workbook.close();
        os.close();
        return outputFile.getAbsolutePath();
    }
    private File getOutputFileName(File excelFileDir) {
        File outputFileName = new File(excelFileDir, OUTPUT_FILE_NAME + ".xlsx");
        int i = 0;
        while(outputFileName.exists()) {
            outputFileName = new File(excelFileDir, OUTPUT_FILE_NAME + "_" + i + ".xlsx");
            i++;
        }
        return outputFileName;
    }
}
