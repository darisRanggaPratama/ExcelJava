package excelMaven;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.FileOutputStream;

public class FontStyles {
    static XSSFWorkbook book;
    static XSSFSheet sheet;
    static XSSFRow row;
    static XSSFFont font;
    static XSSFCellStyle style;
    static XSSFCell cell;
    static FileOutputStream result;
    static File file;
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;



    public static void styledFont(){
        try{

            book = new XSSFWorkbook();
            sheet = book.createSheet("FontStyle");
            row = sheet.createRow(2);

            // Create a new font and alter it
            font = book.createFont();
            font.setFontHeightInPoints((short) 30);
            font.setFontName("IMPACT");
            font.setItalic(true);
            font.setColor(HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.getIndex());

            // Set font style
            style = book.createCellStyle();
            style.setFont(font);

            cell = row.createCell(1);
            cell.setCellValue("Font Styles");
            cell.setCellStyle(style);

            file = new File("./src/main/java/checking/StyleFont.xlsx");
            result = new FileOutputStream(file);
            book.write(result);
            result.close();
            System.out.println("File created successfully");

        } catch (Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
