package bscpoi;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WorkbookCreated {
    public static void main(String[] args) throws IOException {
        // Create blank workbook
        XSSFWorkbook book = new XSSFWorkbook();
        // Create file system
        FileOutputStream outFile = new FileOutputStream(
                new File("./src/main/java/output/blankBook.xlsx")
        );
        // Write to workbook
        book.write(outFile);
        System.out.println("Success to create blank file");
    }
}
