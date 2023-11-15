package basic;
import java.lang.System.Logger;
import java.lang.System.Logger.Level;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class NegCell {
    private static final Logger logger = System.getLogger("NegCell");

    public static void main(String[] args) {
        try {
            // Path file Excel
            String excelFilePath = "./src/main/java/basic/sample.xlsx";

            // Open file Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Get sheet 1 (index 0)
            Sheet sheet = workbook.getSheetAt(0);

            // Get negative cell in column 1 to 6
            logger.log(Level.INFO, "\n\nAngka negatif ditemukan di: ");
            System.out.println("  Col " + " Row");
            int x = 0;
            for (int colIdx = 0; colIdx < 6; colIdx++) {

                for (int rowIdx = 0; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {

                    Row row = sheet.getRow(rowIdx);
                    if (row != null) {

                        Cell cell = row.getCell(colIdx);

                        // Check negative cell
                        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                            double cellValue = cell.getNumericCellValue();
                            if (cellValue < 0) {
                                x++;
                                System.out.println(x + " " + (colIdx + 1) + "    " + (rowIdx + 1) + ": " + cellValue);
                            }
                        }
                    }
                }
            }

            // Close file Excel
            workbook.close();
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
