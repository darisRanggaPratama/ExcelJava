package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Formula {
    static XSSFWorkbook book;
    static XSSFSheet sheet;
    static XSSFRow row;
    static XSSFCell cell;
    static FileOutputStream result;
    static File file;
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static String txtInfo;

    public static void calculate() {
        try {
            book = new XSSFWorkbook();
            sheet = book.createSheet("Formula");
            row = sheet.createRow(1);
            cell = row.createCell(1);
            cell.setCellValue("A      =");
            cell = row.createCell(2);
            cell.setCellValue(2);

            row = sheet.createRow(2);
            cell = row.createCell(1);
            cell.setCellValue("B      =");
            cell = row.createCell(2);
            cell.setCellValue(4);

            row = sheet.createRow(3);
            cell = row.createCell(1);
            cell.setCellValue("Total    =");
            cell = row.createCell(2);

            // Formula: SUM
            cell.setCellFormula("SUM(C2:C3)");
            cell = row.createCell(3);
            cell.setCellValue("SUM(C2:C3)");

            row = sheet.createRow(4);
            cell = row.createCell(1);
            cell.setCellValue("POWER = ");
            cell = row.createCell(2);

            // Formula: POWER
            cell.setCellFormula("POWER(C2,C3)");
            cell = row.createCell(3);
            cell.setCellValue("POWER(C2,C3)");

            row = sheet.createRow(5);
            cell = row.createCell(1);
            cell.setCellValue("MAX = ");
            cell = row.createCell(2);

            // Formula: MAX
            cell.setCellFormula("MAX(C2,C3)");
            cell = row.createCell(3);
            cell.setCellValue("MAX(C2,C3)");

            row = sheet.createRow(6);
            cell = row.createCell(1);
            cell.setCellValue("FACT = ");
            cell = row.createCell(2);

            // Formula: FACT
            cell.setCellFormula("FACT(C3)");
            cell = row.createCell(3);
            cell.setCellValue("FACT(C3)");

            row = sheet.createRow(7);
            cell = row.createCell(1);
            cell.setCellValue("SQRT = ");
            cell = row.createCell(2);

            // Formula: SQRT
            cell.setCellFormula("SQRT(C5)");
            cell = row.createCell(3);
            cell.setCellValue("SQRT(C5)");

            book.getCreationHelper().createFormulaEvaluator().evaluateAll();

            file = new File("./src/main/java/checking/Formula.xlsx");
            result = new FileOutputStream(file);
            book.write(result);
            result.close();
            System.out.println("File created successfully");

        } catch (Exception e) {
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
