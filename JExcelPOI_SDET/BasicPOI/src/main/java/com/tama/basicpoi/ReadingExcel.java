package com.tama.basicpoi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {
    static String excelFilePath;
    static FileInputStream inputStream;
    static XSSFWorkbook book;
    static XSSFSheet sheet;
    static XSSFRow row;
    static XSSFCell cell;
    static int rows, cols;

    public static void readFile() throws FileNotFoundException, IOException {
        excelFilePath = "./src/main/java/com/tama/basicpoi/dataFiles/countries.xlsx";       
        inputStream = new FileInputStream(excelFilePath);

        book = new XSSFWorkbook(inputStream);
        sheet = book.getSheet("data"); // XSSFSheet sheet1 = book.getSheetAt(0); cara lain
        
        // Using Loop        
        rows = sheet.getLastRowNum();
        cols = sheet.getRow(1).getLastCellNum();
        
        for (int r = 0; r <= rows; r++){            
            row = sheet.getRow(r);
            
            for (int c = 0; c < cols; c++){               
                cell = row.getCell(c);
                
                switch(cell.getCellType()){
                    case STRING: System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC: System.out.print((long) cell.getNumericCellValue()); break;
                    case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
                }
                System.out.print("\t\t");
            }
            System.out.println("");
        }        
    }
}
