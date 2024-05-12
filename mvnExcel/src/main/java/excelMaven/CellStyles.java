package excelMaven;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class CellStyles {
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;
    static XSSFWorkbook workbook;
    static XSSFSheet worksheet;
    static XSSFRow row;
    static XSSFCell cell;
    static FileOutputStream result;
    static File file;

    static XSSFCellStyle topLeftAlign, centerAlignContent, bottomRight, justifiedAlign, cellBorder, backColor, foreColor;

    public static void styledCell() {
        try {
            workbook = new XSSFWorkbook();
            worksheet = workbook.createSheet("Cell Style");

            row = worksheet.createRow((short) 1);
            row.setHeight((short) 800);
            cell = row.createCell((short) 1);
            cell.setCellValue("Test of merging");

            // Merge Cells.
            worksheet.addMergedRegion(new CellRangeAddress(
                    1, // first row
                    1, // last row
                    1, // first column
                    4 // last column
            ));

            // Cell Alignment
            row = worksheet.createRow(5);
            cell = row.createCell(0);
            row.setHeight((short) 800);

            // Top Left Alignment
            topLeftAlign = workbook.createCellStyle();
            worksheet.setColumnWidth(0, 8000);
            topLeftAlign.setAlignment(HorizontalAlignment.LEFT);
            topLeftAlign.setVerticalAlignment(VerticalAlignment.TOP);
            cell.setCellValue("TOP LEFT");
            cell.setCellStyle(topLeftAlign);

            row = worksheet.createRow(6);
            cell = row.createCell(1);
            row.setHeight((short) 800);

            // Center Align Cell Contents
            centerAlignContent = workbook.createCellStyle();
            centerAlignContent.setAlignment(HorizontalAlignment.CENTER);
            centerAlignContent.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellValue("Center Aligned");
            cell.setCellStyle(centerAlignContent);

            row = worksheet.createRow(7);
            cell = row.createCell(2);
            row.setHeight((short) 800);

            // Bottom Right Alignment
            bottomRight = workbook.createCellStyle();
            bottomRight.setAlignment(HorizontalAlignment.RIGHT);
            bottomRight.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cell.setCellValue("BOTTOM RIGHT");
            cell.setCellStyle(bottomRight);

            row = worksheet.createRow(8);
            cell = row.createCell(3);

            // Justified Alignment
            justifiedAlign = workbook.createCellStyle();
            justifiedAlign.setAlignment(HorizontalAlignment.JUSTIFY);
            justifiedAlign.setVerticalAlignment(VerticalAlignment.JUSTIFY);
            cell.setCellValue("Contents are justified in alingment");
            cell.setCellStyle(justifiedAlign);

            row = worksheet.createRow((short) 10);
            row.setHeight((short) 800);
            cell = row.createCell((short) 1);

            // Cell Border
            cell.setCellValue("BORDER");
            cellBorder = workbook.createCellStyle();
            cellBorder.setBorderBottom(BorderStyle.THICK);
            cellBorder.setBottomBorderColor(IndexedColors.BLUE.getIndex());
            cellBorder.setBorderLeft(BorderStyle.DOUBLE);
            cellBorder.setLeftBorderColor(IndexedColors.GREEN.getIndex());
            cellBorder.setBorderRight(BorderStyle.HAIR);
            cellBorder.setRightBorderColor(IndexedColors.RED.getIndex());
            cellBorder.setBorderTop(BorderStyle.DOTTED);
            cellBorder.setTopBorderColor(IndexedColors.CORAL.getIndex());
            cell.setCellStyle(cellBorder);

            // Fill Color. Background color
            row = worksheet.createRow((short) 11);
            cell = row.createCell((short) 1);
            backColor = workbook.createCellStyle();
            backColor.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.getIndex());
            backColor.setFillPattern(FillPatternType.LESS_DOTS);
            backColor.setAlignment(HorizontalAlignment.FILL);
            worksheet.setColumnWidth(1, 8000);
            cell.setCellValue("FILL BACKGROUND/ FILL PATTERN");
            cell.setCellStyle(backColor);

            // Fill Color. Foreground color
            row = worksheet.createRow((short) 12);
            cell = row.createCell((short) 1);
            foreColor = workbook.createCellStyle();
            foreColor.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
            foreColor.setFillPattern(FillPatternType.LESS_DOTS);
            foreColor.setAlignment(HorizontalAlignment.FILL);
            cell.setCellValue("FILL FOREGROUND/ FILL PATTERN");
            cell.setCellStyle(foreColor);

            file = new File("./src/main/java/checking/StyleSheet.xlsx");
            result = new FileOutputStream(file);
            workbook.write(result);
            result.close();
            System.out.println("File created successfully");

        } catch (Exception e) {
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
