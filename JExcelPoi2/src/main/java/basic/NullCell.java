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
import java.util.Optional;

public class NullCell {
    private static final System.Logger logger = System.getLogger("NullCell");

    public static void checkNull() {
        try {
            // Get Excel file path
            String path = "./src/main/java/trial/";
            String excelFilePath = path + Data.fileXl;

            // Open file
            FileInputStream input = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(input);

            // Get worksheet (GAJI)
            Sheet sheet = workbook.getSheet(Data.sheetXl);

            // Create new sheet
            Sheet sheetNew = workbook.createSheet("Null");

            // Create Cell Title
            Row rowTitle = sheetNew.createRow(0);
            Cell titleNo = rowTitle.createCell(0);
            titleNo.setCellValue("No");
            Cell cellTitle = rowTitle.createCell(1);
            cellTitle.setCellValue("Cell Null/Empty");
            System.out.println("\nNo  Null/Empty");

            // Check null/empty/missing/blank cell
            int rowIndex = 1, rowIdx = 0, colIdx = 0, no = 0;
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            for (rowIdx = Data.beginRow; rowIdx <= Data.endRow; rowIdx++) {

                int idxRow = Optional.ofNullable(rowIdx)
                        .orElse(0);

                Row row = sheet.getRow(idxRow);

                for (colIdx = Data.firstColumn; colIdx <= Data.lastColumn; colIdx++) {

                    int idxCol = Optional.ofNullable(colIdx)
                            .orElse(0);

                    Cell cell = row.getCell(idxCol);

                    // Check null/empty/missing/blank value
                    if (cell == null || cell.getCellType() == CellType.BLANK ||
                            (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isBlank()) ||
                            (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())) {
                        no++;

                        // Add comment
                        Comment comment = sheet.createDrawingPatriarch().createCellComment(
                                new XSSFClientAnchor(0, 0, 0, 0,
                                        (short) colIdx, rowIdx, (short) (colIdx + 1), rowIdx + 1));
                        comment.setString(new XSSFRichTextString("Null"));
                        cell.setCellComment(comment);

                        // Add Background color
                        CellStyle style = workbook.createCellStyle();
                        style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        cell.setCellStyle(style);

                        // Write blank address cell in new sheet
                        Row newRow = sheetNew.createRow(rowIndex++);
                        Cell newNo = newRow.createCell(0);
                        Cell newCellAddress = newRow.createCell(1);

                        String cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                        newNo.setCellValue(no);
                        newCellAddress.setCellValue(cellAddress);

                        // Write in console
                        System.out.println(no + "   " + cellAddress);
                    }
                }
            }

            // Save file
            String outFile = Data.sheetXl + "_Null.xlsx";
            String outPath = "./src/main/java/output/" + outFile;
            System.out.println("\nOutput: \n" + outPath);


            FileOutputStream output = new FileOutputStream(outPath);
            workbook.write(output);

            // Close Excel file
            workbook.close();
            input.close();
            output.close();
        } catch (IOException e) {
            String noFile = "\nFile not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, noFile);
        } catch (NullPointerException e) {
            String blank = "\nBlank Cell: \n" + e;
            logger.log(System.Logger.Level.ERROR, blank);
        }
    }
}
