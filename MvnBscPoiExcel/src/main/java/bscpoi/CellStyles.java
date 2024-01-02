package bscpoi;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellStyles {
    public static void main(String[] args) throws Exception {
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet("cellStyle");
        XSSFRow row = sheet.createRow((1));
        row.setHeight((short) 800);
        XSSFCell cell = row.createCell( 1);
        cell.setCellValue("TEST OF MERGING");

        // Merging Cells
        sheet.addMergedRegion(new CellRangeAddress(
                1, // First row (0-based)
                1, // Last row (0-based)
                1, // First column (0-based)
                4 // Last column
        ));

        // Cell Alignment
        row = sheet.createRow(5);
        cell = row.createCell(0);
        row.setHeight((short) 800);






    }
}
