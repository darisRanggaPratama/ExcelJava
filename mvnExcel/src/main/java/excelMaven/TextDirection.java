package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.FileOutputStream;

public class TextDirection {
    static XSSFWorkbook wbook;
    static XSSFSheet wsheet;
    static XSSFRow row;
    static XSSFCellStyle styles;
    static XSSFCell cell;
    static FileOutputStream out;
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;
    static File file;

    public static void direction(){
        try {
            wbook = new XSSFWorkbook();
            wsheet = wbook.createSheet("Text direction");
            row = wsheet.createRow(2);
            styles = wbook.createCellStyle();
            styles.setRotation((short) 0);
            cell = row.createCell(1);
            cell.setCellValue("0 Degree angle");
            cell.setCellStyle(styles);

            // 30 Degree
            styles = wbook.createCellStyle();
            styles.setRotation((short) 30);
            cell = row.createCell(3);
            cell.setCellValue("30 Degree angle");
            cell.setCellStyle(styles);

            // 90 Degree
            styles = wbook.createCellStyle();
            styles.setRotation((short) 90);
            cell = row.createCell(5);
            cell.setCellValue("90 Degree angle");
            cell.setCellStyle(styles);

            // 120 Degree
            styles = wbook.createCellStyle();
            styles.setRotation((short) 120);
            cell = row.createCell(7);
            cell.setCellValue("120 Degree angle");
            cell.setCellStyle(styles);

            // 270 Degree
            styles = wbook.createCellStyle();
            styles.setRotation((short) 270);
            cell = row.createCell(9);
            cell.setCellValue("270 Degree angle");
            cell.setCellStyle(styles);

            // 360 Degree
            styles = wbook.createCellStyle();
            styles.setRotation((short) 360);
            cell = row.createCell(12);
            cell.setCellValue("360 Degree angle");
            cell.setCellStyle(styles);

            file = new File("./src/main/java/checking/textDirection.xlsx");
            out = new FileOutputStream(file);
            wbook.write(out);
            out.close();
            System.out.println("File created successfully");

        } catch (Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
