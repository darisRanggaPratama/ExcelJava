package bscpoi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

public class CellDataType {
    public static void main(String[] args) throws Exception {
        // Create workbook
        XSSFWorkbook book = new XSSFWorkbook();
        // Create worksheet
        XSSFSheet sheet = book.createSheet("Data types");
        // Create row
        XSSFRow row = sheet.createRow(2);

        row.createCell(0).setCellValue("Data Type");
        row.createCell(1).setCellValue("Cell value");

        row = sheet.createRow(3);
        row.createCell(0).setCellValue("Type: Blank");
        row.createCell(1);

        row = sheet.createRow(4);
        row.createCell(0).setCellValue("Type: Boolean");
        row.createCell(1).setCellValue(true);

        row = sheet.createRow(5);
        row.createCell(0).setCellValue("Type: Error");
        row.createCell(1).setCellValue(String.valueOf(CellType.ERROR));

        row = sheet.createRow(6);
        row.createCell(0).setCellValue("Type: Date");
        row.createCell(1).setCellValue(new Date());

        row = sheet.createRow(7);
        row.createCell(0).setCellValue("Type: Numeric");
        row.createCell(1).setCellValue(20);

        row = sheet.createRow(8);
        row.createCell(0).setCellValue("Type: String");
        row.createCell(1).setCellValue("Text");

        FileOutputStream outFile = new FileOutputStream(new File("./src/main/java/output/dataType.xlsx"));
        book.write(outFile);
        outFile.close();
        System.out.println("Create workbook");

    }
}
