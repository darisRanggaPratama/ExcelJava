package kotlins

import basic.Data
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

import java.lang.System.getLogger
import java.lang.System.Logger.Level

class NegativeCellK {
    companion object {
        private val logger = getLogger("NegativeCellK")

        fun checkNegative() {
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
                val workbook = XSSFWorkbook(inputStream)
                val sheet = workbook.getSheet(Data.sheetXl)

                val newSheet = workbook.createSheet("NEG")
                println("$MAGENTA\nFound$RESET Negative: ")
                println(YELLOW + " Cell " + RESET + CYAN + "Value" + RESET)

                val titleRow = newSheet.createRow(0)
                val titleNo = titleRow.createCell(0)
                titleNo.setCellValue("No")
                val titleCell = titleRow.createCell(1)
                titleCell.setCellValue("Cell")
                val titleValue = titleRow.createCell(2)
                titleValue.setCellValue("Negative")

                var rowIndex: Int = 1
                var no: Int = 0
                for (rowIdx in Data.beginRow..Data.endRow) {
                    val row = sheet.getRow(rowIdx)
                    for (colIdx in Data.firstColumn..Data.lastColumn) {
                        val cell = row.getCell(colIdx)

                        if (cell != null && cell.cellType == CellType.NUMERIC) {
                            val cellValue = cell.numericCellValue

                            if (cellValue < 0) {

                                val comment = sheet.createDrawingPatriarch().createCellComment(
                                    XSSFClientAnchor(
                                        0, 0, 0, 0, colIdx.toInt(), rowIdx, (colIdx + 1).toInt(), (rowIdx + 1)
                                    )
                                )
                                comment.string = XSSFRichTextString("Negative: $cellValue")
                                cell.cellComment = comment

                                val style = workbook.createCellStyle()
                                style.fillForegroundColor = IndexedColors.YELLOW.index
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

                val outputStream = FileOutputStream("./src/main/java/output/KNeg.xlsx")
                workbook.write(outputStream)

                workbook.close()
                inputStream.close()
                outputStream.close()

            } catch (e: IOException) {
                val noFile = "File not found: \n$e"
                logger.log(Level.ERROR, noFile)

            } catch (e: NullPointerException) {
                val blank = "Blank cell: \n$e"
                logger.log(Level.ERROR, blank)
            }
        }
    }
}