package excelMaven;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

public class DataTypes {
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static XSSFWorkbook workbook;
    static XSSFSheet worksheet;
    static XSSFRow row;
    static FileOutputStream result;
    static File file;
    static String txtInfo;

    public static void excelData(){
        try{
            workbook = new XSSFWorkbook();
            worksheet = workbook.createSheet("Data Type");

            row = worksheet.createRow((short) 2);
            row.createCell(0).setCellValue("Type of Cell");
            row.createCell(1).setCellValue("Cell value");

            row = worksheet.createRow((short) 3);
            row.createCell(0).setCellValue("Set Cell type: BLANK");
            row.createCell(1);

            row = worksheet.createRow((short) 4);
            row.createCell(0).setCellValue("Set Cell type: BOOLEAN");
            row.createCell(1).setCellValue(true);

            row = worksheet.createRow((short) 5);
            row.createCell(0).setCellValue("Set Cell type: ERROR");
            row.createCell(1).setCellValue(String.valueOf(CellType.ERROR));

            row = worksheet.createRow((short) 6);
            row.createCell(0).setCellValue("Set Cell type: DATE");
            row.createCell(1).setCellValue(new Date());

            row = worksheet.createRow((short) 7);
            row.createCell(0).setCellValue("Set Cell type: Numeric");
            row.createCell(1).setCellValue(1924);

            row = worksheet.createRow((short) 8);
            row.createCell(0).setCellValue("Set Cell type: String");
            row.createCell(1).setCellValue("Rangga");

            file = new File("./src/main/java/checking/DataType.xlsx");
            result = new FileOutputStream(file);
            workbook.write(result);
            result.close();
            System.out.println("File created successfully");

        } catch (Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
