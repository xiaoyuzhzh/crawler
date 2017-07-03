package com.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Date;

/**
 * Created by xiaoy on 2017/6/13.
 */
public class ExcelFileGenerator {

    public void saveAsExcelFile() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("d:/workbook.xlsx");


        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])
        Sheet sheet1 = wb.createSheet("sheet1");

        Row row = sheet1.createRow(0);
        row.createCell(0).setCellValue(1);
        row.createCell(1).setCellValue(2);
        row.createCell(2).setCellValue(true);
        row.createCell(3).setCellValue(new Date());
        row.createCell(4).setCellValue(111);



        CreationHelper creationHelper = wb.getCreationHelper();
        Row row2 = sheet1.createRow(1);
        XSSFCellStyle cellStyle2 = wb.createCellStyle();
        cellStyle2.setFillForegroundColor(new XSSFColor(new java.awt.Color(153,204,255)));
        cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Cell cell2 = row2.createCell(2);
        cell2.setCellStyle(cellStyle2);
        cell2.setCellValue(creationHelper.createRichTextString("aaa\nbbb"));

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                creationHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        Cell cell = row2.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        wb.write(fileOut);
        fileOut.close();

    }

    public static void main(String[] args) throws Exception {
        ExcelFileGenerator excelFileGenerator = new ExcelFileGenerator();
        excelFileGenerator.saveAsExcelFile();
    }
}
