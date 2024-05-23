package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class PrintArea {
    static XSSFWorkbook book;
    static XSSFSheet sheet;
    static File file;
    static FileOutputStream output;
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;
    public static void printing(){
        try {
            book = new XSSFWorkbook();
            sheet = book.createSheet("Print Area");

            // Set Print Area with indexes
            book.setPrintArea(
                    0, // Sheet Index
                    0, // Start Column
                    5, // End Column
                    0, // Start Row
                    5 // End Row
            );

            // Set Paper Size
            sheet.getPrintSetup().setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);

            // Set Display Grid Lines or not
            sheet.setDisplayGridlines(true);

            // Set Print Grid Lines or not
            sheet.setPrintGridlines(true);

            file = new File("./src/main/java/checking/PrintArea.xlsx");
            output = new FileOutputStream(file);
            book.write(output);
            output.close();
            System.out.println("File created successfully");
        } catch (Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
