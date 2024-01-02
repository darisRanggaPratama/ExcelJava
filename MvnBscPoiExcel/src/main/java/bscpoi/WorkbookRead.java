package bscpoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class WorkbookRead {
    static XSSFRow row;

    public static void main(String[] args) throws Exception {
        FileInputStream inputFile = new FileInputStream(new File("./src/main/java/trial/sample.xlsx"));

        XSSFWorkbook book = new XSSFWorkbook(inputFile);
        XSSFSheet sheet = book.getSheetAt(0);
        Iterator<Row> rowIterat = sheet.iterator();

        while (rowIterat.hasNext()) {
            row = (XSSFRow) rowIterat.next();
            Iterator<Cell> cellIterat = row.cellIterator();
            while (cellIterat.hasNext()) {
                Cell cell = cellIterat.next();
                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " \t\t ");
                        break;

                    case STRING:
                        System.out.print(cell.getStringCellValue() + " \t\t ");
                        break;
                }
            }
            System.out.println();
        }
        inputFile.close();
    }
}

