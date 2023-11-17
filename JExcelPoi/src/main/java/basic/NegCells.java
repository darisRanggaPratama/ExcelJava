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

public class NegCells {
    public static void main(String[] args) {
        try {
            // Menentukan path file Excel
            String excelFilePath = "./src/main/java/trial/sample.xlsx";

            // Membuka file Excel
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            // Mendapatkan sheet dengan nama "GAJI"
            Sheet sheetGAJI = workbook.getSheet("GAJI");

            // Membuat sheet baru dengan nama "Neg" untuk menampilkan hasil
            Sheet sheetNeg = workbook.createSheet("Neg");

            // Membuat baris pertama di sheet "Neg" untuk judul
            Row negHeaderRow = sheetNeg.createRow(0);
            Cell negCellAddressTitle = negHeaderRow.createCell(0);
            negCellAddressTitle.setCellValue("Alamat Cell Negatif");
            Cell negCellValueTitle = negHeaderRow.createCell(1);
            negCellValueTitle.setCellValue("Nilai Negatif");

            // Mengecek dan menandai nilai negatif dengan warna kuning
            int negRowIndex = 1;
            int y = 0;
            for (int rowIdx = 1; rowIdx <= 45; rowIdx++) {
                Row row = sheetGAJI.getRow(rowIdx);

                for (int colIdx = 0; colIdx <= 4; colIdx++) {
                    Cell cell = row.getCell(colIdx);

                    // Mengecek apakah nilai di sel adalah angka negatif
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        double cellValue = cell.getNumericCellValue();
                        if (cellValue < 0) {
                            y++;

                            Comment comment = sheetGAJI.createDrawingPatriarch().createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) colIdx, rowIdx, (short) (colIdx + 1), rowIdx + 1));
                            comment.setString(new XSSFRichTextString("Negative: " + cellValue));
                            cell.setCellComment(comment);


                            // Menandai nilai negatif dengan warna kuning di sheet "GAJI"
                            CellStyle style = workbook.createCellStyle();
                            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell.setCellStyle(style);

                            // Menampilkan alamat sel negatif beserta nilainya di sheet "Neg"
                            Row negRow = sheetNeg.createRow(negRowIndex++);
                            Cell negCellAddress = negRow.createCell(0);
                            negCellAddress.setCellValue(CellReference.convertNumToColString(colIdx) + (rowIdx + 1));

                            Cell negCellValue = negRow.createCell(1);
                            negCellValue.setCellValue(cellValue);

                            System.out.println(y + " " + CellReference.convertNumToColString(colIdx) + (rowIdx + 1) + ": " + cellValue);
                        }
                    }
                }
            }

            // Menyimpan perubahan ke file Excel
            FileOutputStream outputStream = new FileOutputStream("./src/main/java/output/Neg.xlsx");
            workbook.write(outputStream);
            workbook.close();
            inputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
