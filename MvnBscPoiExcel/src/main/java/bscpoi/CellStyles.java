package bscpoi;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.FileOutputStream;
import java.io.File;


public class CellStyles {
    public static void main(String[] args) throws Exception {
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("cellStyle");
        XSSFRow row = sheet.createRow((1));
        row.setHeight((short) 800);
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("TEST OF MERGING");

        // Merging Cells
        sheet.addMergedRegion(new CellRangeAddress(
                1, // First row (0-based)
                1, // Last row (0-based)
                1, // First column (0-based)
                4 // Last column
        ));

        // Cell Alignment
        row = sheet.createRow(5);
        cell = row.createCell(0);
        row.setHeight((short) 800);

        // Top Left alignment
        XSSFCellStyle style1 = book.createCellStyle();
        sheet.setColumnWidth(0, 8000);
        style1.setAlignment(HorizontalAlignment.LEFT);
        style1.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellValue("TOP LEFT");
        cell.setCellStyle(style1);

        row = sheet.createRow(6);
        cell = (XSSFCell) row.createCell(1);
        row.setHeight((short) 800);

        // Cell Align Cell Contents
        XSSFCellStyle style2 = book.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellValue("CENTER ALIGNED");
        cell.setCellStyle(style2);

        row = sheet.createRow(7);
        cell = (XSSFCell) row.createCell(2);
        row.setHeight((short) 800);

        // Bottom Right alignment
        XSSFCellStyle style3 = book.createCellStyle();
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellValue("BOTTOM RIGHT");
        cell.setCellStyle(style3);

        row = sheet.createRow(8);
        cell = (XSSFCell) row.createCell(3);

        // Justified Alignment
        XSSFCellStyle style4 = book.createCellStyle();
        style4.setAlignment(HorizontalAlignment.JUSTIFY);
        style4.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        cell.setCellValue("CONTENTS ARE JUSTIFIED IN ALIGNMENT");
        cell.setCellStyle(style4);

        // CELL BORDER
        row = sheet.createRow((short) 10);
        row.setHeight((short) 800);
        cell = (XSSFCell) row.createCell((short) 1);
        XSSFCellStyle style5 = book.createCellStyle();
        style5.setBorderBottom(BorderStyle.THICK);
        style5.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        style5.setBorderLeft(BorderStyle.DOUBLE);
        style5.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style5.setBorderLeft(BorderStyle.HAIR);
        style5.setRightBorderColor(IndexedColors.RED.getIndex());
        style5.setBorderTop(BorderStyle.DOTTED);
        style5.setTopBorderColor(IndexedColors.CORAL.getIndex());
        cell.setCellValue("BORDER");
        cell.setCellStyle(style5);

        // Fill Colors
        // Background Color
        row = sheet.createRow((short) 11);
        cell = (XSSFCell) row.createCell((short) 1);
        XSSFCellStyle style6 = book.createCellStyle();
        style6.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.getIndex());
        style6.setFillPattern(FillPatternType.LESS_DOTS);
        style6.setAlignment(HorizontalAlignment.FILL);
        sheet.setColumnWidth(1, 8000);
        cell.setCellValue("FILL BACKGROUND/ FILL PATTERN");
        cell.setCellStyle(style6);

        // Foreground Color
        row = sheet.createRow((short) 12);
        cell = (XSSFCell) row.createCell((short) 1);
        XSSFCellStyle style7 = book.createCellStyle();
        style7.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        style7.setFillPattern(FillPatternType.LESS_DOTS);
        style7.setAlignment(HorizontalAlignment.FILL);
        cell.setCellValue("FILL FOREGROUND/ FILL PATTERN");
        cell.setCellStyle(style7);

        FileOutputStream out = new FileOutputStream(new File("./src/main/java/output/CellStyle.xlsx"));
        book.write(out);
        out.close();
        System.out.println("File written successfully");


    }}