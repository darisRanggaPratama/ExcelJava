package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.lang.System.Logger;

public class CheckExistFile {
    static Logger logger = System.getLogger("CheckWorkbook");
    static File file;
    static FileInputStream entry;

    static XSSFWorkbook workbook;

    public static void openExistWorkbook() throws IOException {
        try{
            file = new File("./src/main/java/checking/BlankWorkbook.xlsx");
            entry = new FileInputStream(file);
        } catch (FileNotFoundException e){
            String noFile = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, noFile);
        } finally {
            workbook = new XSSFWorkbook(entry);
            if (file.isFile() && file.exists()){
                String noFile = "File open successfully: \n";
                logger.log(System.Logger.Level.ERROR, noFile);
            }
        }
    }
}
