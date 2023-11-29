package basic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Comment;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.System.Logger;
import java.lang.System.Logger.Level;

public class ErrCell {
    private static final Logger logger = System.getLogger("ErrCell");

    public static void checkErrCell() {
        // COLOR
        String BLUE = "\u001B[34m", RESET = "\u001B[0m", RED = "\u001B[31m";
        String MAGENTA = "\u001B[45m", YELLOW = "\u001B[43m", CYAN = "\u001B[46m";

        try {
            // Get Excel file path
            String name = Data.fileXl;
            String path = "./src/main/java/trial/";
            String excelFilePath = path + name;

            // Open file
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Get worksheet (GAJI)
            String sheetName = Data.sheetXl;
            Sheet sheet = workbook.getSheet(sheetName);

            // Create worksheet
            Sheet newSheet = workbook.createSheet("Err");

            // Get Decimal. Example: A2 to E46
            int firstRow = Data.beginRow;
            int lastRow = Data.endRow;
            int firstCol = Data.firstColumn;
            int lastCol = Data.lastColumn;

            System.out.println(MAGENTA + "\nFound" + RESET + " ERROR: ");
            System.out.println(YELLOW + " Cell " + RESET + CYAN + "Value" + RESET);

            // Cell to place title
            Row titleRow = newSheet.createRow(0);
            Cell titleNo = titleRow.createCell(0);
            titleNo.setCellValue("No");
            Cell titleCell = titleRow.createCell(1);
            titleCell.setCellValue("Cell");
            Cell titleValue = titleRow.createCell(2);
            titleValue.setCellValue("Error");

            int rowIndex = 1, no = 0;
            for (int rowIdx = firstRow; rowIdx <= lastRow; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
                    Cell cell = row.getCell(colIdx);

                    // Check Decimal Value
                    if (cell != null && cell.getCellType() == CellType.FORMULA) {
                        String cellValue = cell.getCellFormula();

                        if (cell.getCachedFormulaResultType() == CellType.ERROR) {

                            Comment comment = sheet.createDrawingPatriarch().createCellComment(
                                    new XSSFClientAnchor(0, 0, 0, 0,
                                            (short) colIdx, rowIdx, (short) (colIdx + 1), rowIdx + 1));
                            comment.setString(new XSSFRichTextString("Error: " + cellValue));
                            cell.setCellComment(comment);

                            CellStyle style = workbook.createCellStyle();
                            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell.setCellStyle(style);

                            Row newRow = newSheet.createRow(rowIndex++);
                            Cell newNo = newRow.createCell(0);
                            Cell newCellAddress = newRow.createCell(1);
                            Cell newCellValue = newRow.createCell(2);

                            String cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                            no++;
                            newNo.setCellValue(no);
                            newCellAddress.setCellValue(cellAddress);
                            newCellValue.setCellValue(cellValue);

                            System.out.println(no + RESET + " " + BLUE + cellAddress + ": " + RESET + RED + cellValue + RESET);
                        }
                    }
                }
            }

            // Save excel file
            FileOutputStream outputStream = new FileOutputStream("./src/main/java/output/Err.xlsx");
            workbook.write(outputStream);

            // Close excel file
            workbook.close();
            inputStream.close();
            outputStream.close();
        } catch (IOException e) {
            String noFile = "File not found: \n" + e;
            logger.log(Level.ERROR, noFile);
        } catch (NullPointerException e) {
            String blank = "Blank Cell: \n" + e;
            logger.log(Level.ERROR, blank);
        }
    }
}
