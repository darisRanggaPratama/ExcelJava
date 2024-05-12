package excelMaven;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadSpreadsheet {
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static File file;
    static FileInputStream entry;
    static XSSFWorkbook workbook;
    static XSSFSheet worksheet;
    static XSSFRow row;
    static Cell cell;
    static String txtInfo;
    public static void readWorksheet(){
        try {
            file = new File("./src/main/java/checking/SheetCreated.xlsx");

            entry = new FileInputStream(file);
            workbook = new XSSFWorkbook(entry);
            worksheet = workbook.getSheet("Person Data");
            Iterator<Row> rowIterator = worksheet.iterator();

            while (rowIterator.hasNext()){
                row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()){
                    cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case CellType.STRING:
                            System.out.print(cell.getStringCellValue() + " \t\t ");
                            break;
                        default:
                            System.out.print(cell.getNumericCellValue() + " \t\t ");
                            break;
                    }
                }
                System.out.println();
            }
            entry.close();
            txtInfo = "File read successfully";
            logger.log(System.Logger.Level.ERROR, txtInfo);

        } catch (Exception e){
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
