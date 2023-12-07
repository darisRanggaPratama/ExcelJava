import basic.Data
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

object StrCell {
    private val logger = System.getLogger("StrCell")
    fun checkString() {
        // COLOR
        val BLUE = "\u001B[34m"
        val RESET = "\u001B[0m"
        val RED = "\u001B[31m"
        val MAGENTA = "\u001B[45m"
        val YELLOW = "\u001B[43m"
        val CYAN = "\u001B[46m"
        try {
            // Get Excel file path
            val path = "./src/main/java/trial/"
            val excelFilePath = path + Data.fileXl

            // Open file
            val inputStream = FileInputStream(excelFilePath)
            val workbook: Workbook = XSSFWorkbook(inputStream)

            // Get worksheet (GAJI)
            val sheet = workbook.getSheet(Data.sheetXl)

            // Create worksheet
            val newSheet = workbook.createSheet("Err")
            println("$MAGENTA\nFound$RESET STRING: ")
            println(YELLOW + " Cell " + RESET + CYAN + "Value" + RESET)

            // Cell to place title
            val titleRow = newSheet.createRow(0)
            val titleNo = titleRow.createCell(0)
            titleNo.setCellValue("No")
            val titleCell = titleRow.createCell(1)
            titleCell.setCellValue("Cell")
            val titleValue = titleRow.createCell(2)
            titleValue.setCellValue("String")
            var rowIndex = 1
            var no = 0
            for (rowIdx in Data.beginRow..Data.endRow) {
                val row = sheet.getRow(rowIdx)
                for (colIdx in Data.firstColumn..Data.lastColumn) {
                    val cell = row.getCell(colIdx)

                    // Check Decimal Value
                    if (cell != null && cell.cellType == CellType.STRING) {
                        val cellValue = cell.stringCellValue
                        val comment = sheet.createDrawingPatriarch().createCellComment(
                            XSSFClientAnchor(
                                0, 0, 0, 0,
                                colIdx.toShort().toInt(), rowIdx, (colIdx + 1).toShort().toInt(), rowIdx + 1
                            )
                        )
                        comment.string = XSSFRichTextString("String: $cellValue")
                        cell.cellComment = comment

                        val style = workbook.createCellStyle()
                        style.fillForegroundColor = IndexedColors.LIGHT_GREEN.getIndex()
                        style.fillPattern = FillPatternType.SOLID_FOREGROUND
                        cell.cellStyle = style
                        val newRow = newSheet.createRow(rowIndex++)
                        val newNo = newRow.createCell(0)
                        val newCellAddress = newRow.createCell(1)
                        val newCellValue = newRow.createCell(2)
                        val cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1)
                        no++
                        newNo.setCellValue(no.toDouble())
                        newCellAddress.setCellValue(cellAddress)
                        newCellValue.setCellValue(cellValue)
                        println("$no$RESET $BLUE$cellAddress: $RESET$RED$cellValue$RESET")
                    }
                }
            }

            // Save excel file
            val outputStream = FileOutputStream("./src/main/java/output/Str.xlsx")
            workbook.write(outputStream)

            // Close excel file
            workbook.close()
            inputStream.close()
            outputStream.close()
        } catch (e: IOException) {
            val noFile = "File not found: \n$e"
            logger.log(System.Logger.Level.ERROR, noFile)
        } catch (e: NullPointerException) {
            val blank = "Blank Cell: \n$e"
            logger.log(System.Logger.Level.ERROR, blank)
        }
    }
}
