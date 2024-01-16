package kotlins

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
import java.lang.System.Logger.Level
import java.lang.System.getLogger

class ZeroCells {
    companion object {

        private val logger = getLogger("ZeroCells")

        fun checkZero() {
            val BLUE = "\u001B[34m"
            val RESET = "\u001B[0m"
            val RED = "\u001B[31m"
            val MAGENTA = "\u001B[45m"
            val YELLOW = "\u001B[43m"
            val CYAN = "\u001B[46m"

            try {
                val path = "./src/main/java/trial/"
                val excelFilePath = path + Data.fileXl

                val inputStream = FileInputStream(excelFilePath)
                val workbook: Workbook = XSSFWorkbook(inputStream)

                val sheet = workbook.getSheet(Data.sheetXl)

                val newSheet = workbook.createSheet("Zero")
                println("$MAGENTA\nFound$RESET Zero: ")
                println(YELLOW + " Cell " + RESET + CYAN + "Value" + RESET)

                val titleRow = newSheet.createRow(0)
                val titleNo = titleRow.createCell(0)
                titleNo.setCellValue("No")
                val titleCell = titleRow.createCell(1)
                titleCell.setCellValue("Cell")
                val titleValue = titleRow.createCell(2)
                titleValue.setCellValue("Zero")

                var rowIndex = 1
                var no = 0

                for (rowIdx in Data.beginRow..Data.endRow) {
                    val row = sheet.getRow(rowIdx)

                    for (colIdx in Data.firstColumn..Data.lastColumn) {
                        val cell = row.getCell(colIdx)

                        if (cell != null && cell.cellType == CellType.NUMERIC) {
                            val cellValue = cell.numericCellValue

                            if (cellValue == 0.0) {

                                val comment = sheet.createDrawingPatriarch().createCellComment(
                                    XSSFClientAnchor(
                                        0, 0, 0, 0, colIdx.toInt(), rowIdx, (colIdx + 1).toInt(), (rowIdx + 1)
                                    )
                                )
                                comment.string = XSSFRichTextString("Zero: $cellValue")
                                cell.cellComment = comment

                                val style = workbook.createCellStyle()
                                style.fillForegroundColor = IndexedColors.LIGHT_BLUE.index
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
                }

                val outputStream = FileOutputStream("./src/main/java/output/KZero.xlsx")
                workbook.write(outputStream)

                workbook.close()
                inputStream.close()
                outputStream.close()
            } catch (e: IOException) {
                val noFile = "File not found: \n$e"
                logger.log(Level.ERROR, noFile)
            } catch (e: NullPointerException) {
                val blank = "Blank Cell: \n$e"
                logger.log(Level.ERROR, blank)
            }
        }
    }
}
