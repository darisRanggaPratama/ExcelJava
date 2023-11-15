package basic;

import java.io.IOException;
import java.lang.System.Logger;
import java.lang.System.Logger.Level;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.util.CellReference;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class NegVal {
    private static final Logger logger = System.getLogger("NegVal");

    public static void main(String[] args) {
        try {
            // Menentukan path file Excel
            String excelFilePath = "./src/main/java/basic/sample.xlsx";

            // Membuka file Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Mendapatkan sheet pertama (indeks 0)
            Sheet sheet = workbook.getSheetAt(0);

            // Mengecek angka negatif dalam rentang kolom 1-6, baris 2-45
            logger.log(Level.INFO, "\n\nAngka negatif ditemukan di: ");
            int x = 0;
            for (int rowIdx = 1; rowIdx <= 45; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                for (int colIdx = 0; colIdx < 6; colIdx++) {
                    Cell cell = row.getCell(colIdx);

                    // Mengecek apakah nilai di sel adalah angka negatif
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        double cellValue = cell.getNumericCellValue();
                        if (cellValue < 0) {
                            String cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                            x++;
                            System.out.println(x + " " + cellAddress + ": " + cellValue);
                        }
                    }
                }
            }

            // Menutup file Excel
            workbook.close();
            inputStream.close();
        } catch (IOException e) {
            String noFile = "File not found: \n" + e;
            logger.log(Level.ERROR, noFile);

        } catch (NullPointerException e) {
            String blank = "Blank Cell: \n" + e;
            logger.log(Level.ERROR, blank);
        }
    }
}
