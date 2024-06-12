package com.tama.basicpoi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

    public static void createFile() throws FileNotFoundException, IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("EmpInfo");

        Object emppData[][] = {
            {"ID", "Name", "Job"},
            {111, "Eren", "Programmer"},
            {112, "Armin", "Manager"},
            {113, "Mikasa", "Doctor"}
        };

        // Using loop
        int rows, cols;
        rows = emppData.length;
        cols = emppData[0].length;

        System.out.println("Baris: " + rows + " Kolom: " + cols);

        for (int r = 0; r < rows; r++) {
            XSSFRow row = sheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.createCell(c);
                Object value = emppData[r][c];

                if (value instanceof String) {
                    cell.setCellValue((String) value);
                }
                if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                }
                if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                }

            }
        }
        
        String output =  "./src/main/java/com/tama/basicpoi/dataFiles/job.xlsx";
        FileOutputStream outStream = new FileOutputStream(output);
        book.write(outStream);
        outStream.close();
        
        System.out.println("Excel File created");
    }
}
