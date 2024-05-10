package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;

import java.lang.System.Logger;

public class CreateBlankWorkbook {
    static Logger logger = System.getLogger("BlankWorkbook");
    static XSSFWorkbook workbook;
    static File file;
    static FileOutputStream out = null;

    public static void blankWorkbook() throws IOException {

        try {
            file = new File(
                    "./src/main/java/checking/BlankWorkbook.xlsx");
            out = new FileOutputStream(file);
        }catch (FileNotFoundException e){
            String noFile = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, noFile);
        }finally {
            workbook = new XSSFWorkbook();
            workbook.write(out);
            assert out != null;
            out.close();
            logger.log(System.Logger.Level.INFO, "File written successfully");
        }
    }
}
