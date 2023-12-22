package bscpoi;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class WorkbookOpen {
    public static void main(String[] args) throws Exception {
        File file = new File("./src/main/java/trial/sample.xlsx");
        FileInputStream inputFile = new FileInputStream(file);

        // Get xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(inputFile);

        if (file.isFile() && file.exists()){
            System.out.println("Open file: Success");
        }
         else {
            System.out.println("Open file: Error");
        }
    }
}
