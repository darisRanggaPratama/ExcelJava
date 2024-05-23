package excelMaven;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;

import java.io.File;
import java.io.FileOutputStream;

public class Hyperlink {
    static XSSFWorkbook book;
    static XSSFSheet sheet;
    static XSSFCell cell;
    static CreationHelper helper;
    static XSSFCellStyle hlinkStyle;
    static XSSFFont hlinkFont;
    static XSSFHyperlink linked;
    static FileOutputStream result;
    static File file;
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;

    public static void linkCreated(){
        try{
            book = new XSSFWorkbook();
            sheet = book.createSheet("Hyperlink");
            helper = book.getCreationHelper();
            hlinkStyle = book.createCellStyle();
            hlinkFont = book.createFont();
            hlinkFont.setUnderline(Font.U_SINGLE);
            hlinkFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
            hlinkStyle.setFont(hlinkFont);

            // URL Link
            cell = sheet.createRow(1).createCell((short) 1);
            cell.setCellValue("URL Link");
            linked = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.URL);
            linked.setAddress("https://www.youtube.com/");
            cell.setHyperlink(linked);
            cell.setCellStyle(hlinkStyle);

            // Linked to a file in current folder
            cell = sheet.createRow(2).createCell((short) 1);
            cell.setCellValue("File Link");
            linked = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.FILE);
            linked.setAddress("Formula.xlsx");
            cell.setHyperlink(linked);
            cell.setCellStyle(hlinkStyle);

            // E-Mail Link
            cell = sheet.createRow(3).createCell((short) 1);
            cell.setCellValue("Email link");
            linked = (XSSFHyperlink) helper.createHyperlink(HyperlinkType.EMAIL);
            linked.setAddress("mailto:rangga.hrdme@gmail.com?"+"subject=Hyperlink");
            cell.setHyperlink(linked);
            cell.setCellStyle(hlinkStyle);

            file = new File("./src/main/java/checking/Hyperlink.xlsx");
            result = new FileOutputStream(file);
            book.write(result);
            result.close();
            System.out.println("File created successfully");

        } catch(Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }

    }
}
