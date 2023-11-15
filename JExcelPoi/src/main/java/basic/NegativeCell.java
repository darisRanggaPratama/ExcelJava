package basic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;
import java.lang.System.Logger;
import java.lang.System.Logger.Level;

public class NegativeCell {
    private static final Logger logger = System.getLogger("NegativeCell");
    private static final Scanner scan = new Scanner(System.in);

    public static void main(String[] args) {
        String BLUE = "\u001B[34m";
        String RESET = "\u001B[0m";
        String RED = "\u001B[31m";
        String MAGENTA = "\u001B[45m";
        String YELLOW = "\u001B[43m";
        String CYAN = "\u001B[46m";

        try {
            // Menentukan path file Excel
            System.out.println("File Name:");
            String name = scan.nextLine();
            String path = "./src/main/java/trial/";
            String excelFilePath = path + name;

            // Membuka file Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Mendapatkan sheet pertama (GAJI)
            System.out.println("Sheet Name:");
            String sheetName = scan.nextLine();
            Sheet sheet = workbook.getSheet(sheetName);

            // Mengecek angka negatif dalam rentang kolom 1-6, baris 2-45
            int x = 0;

            System.out.print("Row:\n  First: ");
            int firstRow = scan.nextInt();
            System.out.print("  Last: ");
            int lastRow = scan.nextInt();

            System.out.print("Column:\n  First: ");
            int firstCol = scan.nextInt();
            System.out.print("  Last: ");
            int lastCol = scan.nextInt();

            System.out.println(MAGENTA + "Found" + RESET + " Negative: ");
            System.out.println(YELLOW + " Cell " + RESET + CYAN + "Value" + RESET);
            // Example: A2 to E46
            for (int rowIdx = firstRow; rowIdx <= lastRow; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
                    Cell cell = row.getCell(colIdx);

                    // Mengecek apakah nilai di sel adalah angka negatif
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        double cellValue = cell.getNumericCellValue();
                        if (cellValue < 0) {
                            String cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                            x++;
                            System.out.println(x + RESET + " " + BLUE + cellAddress + ": " + RESET + RED + cellValue + RESET);
                        }
                    }
                }
            }

            // Menutup file Excel
            workbook.close();
            scan.close();
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
