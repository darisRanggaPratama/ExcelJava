package kotlins

import basic.Data
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

import java.lang.System.getLogger
import java.lang.System.Logger.Level
import java.util.Optional

class NullCells {
    companion object {
        private val logger = getLogger("NullCells")

        fun checkNull() {
            try {
                val path = "./src/main/java/trial/"
                val excelFilePath = path + Data.fileXl

                val input = FileInputStream(excelFilePath)
                val workbook = XSSFWorkbook(input)
                workbook.missingCellPolicy = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK
                val sheet = workbook.getSheet(Data.sheetXl)
                val sheetNew = workbook.createSheet("Null")

                val rowTitle = sheetNew.createRow(0)
                val titleNo = rowTitle.createCell(0)
                titleNo.setCellValue("No")
                val cellTitle = rowTitle.createCell(1)
                cellTitle.setCellValue("Cell Null/Empty")
                println("\nNo Null/Empty")

                var rowIndex: Int = 1
                var rowIdx: Int = 0
                var colIdx: Int = 0
                var no: Int = 0

                for (rowIdx in Data.beginRow..Data.endRow) {
                    val idxRow = Optional.ofNullable(rowIdx).orElse(0)
                    val row = sheet.getRow(idxRow)

                    for (colIdx in Data.firstColumn..Data.lastColumn) {
                        val idxCol = Optional.ofNullable(colIdx).orElse(0)
                        val cell = row.getCell(idxCol)

                        if (cell == null ||
                            cell.cellType == CellType.BLANK ||
                            cell.cellType == CellType.STRING && cell.stringCellValue.trim().isBlank() ||
                            cell.cellType == CellType.STRING && cell.stringCellValue.trim().isEmpty()
                        ) {

                            no++

                            val comment = sheet.createDrawingPatriarch().createCellComment(
                                XSSFClientAnchor(
                                    0, 0, 0, 0, colIdx.toInt(), rowIdx, (colIdx + 1).toInt(), (rowIdx + 1)
                                )
                            )
                            comment.string = XSSFRichTextString("Null")
                            cell!!.cellComment = comment

                            val style = workbook.createCellStyle()
                            style.fillForegroundColor = IndexedColors.LIGHT_ORANGE.index
                            style.fillPattern = FillPatternType.SOLID_FOREGROUND
                            cell.cellStyle = style

                            val newRow = sheetNew.createRow(rowIndex++)
                            val newNo = newRow.createCell(0)
                            val newCellAddress = newRow.createCell(1)
                            val cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1)
                            newNo.setCellValue(no.toDouble())
                            newCellAddress.setCellValue(cellAddress)

                            println("$no   $cellAddress")
                        }
                    }
                }

                val output = FileOutputStream("./src/main/java/output/KNull.xlsx")
                workbook.write(output)

                workbook.close()
                input.close()
                output.close()
            } catch (e: IOException) {
                val noFile = "\nFile not found: \n$e"
                logger.log(Level.ERROR, noFile)
            } catch (e: NullPointerException) {
                val blank = "\nBlank cell: \n$e"
                logger.log(Level.ERROR, blank)
            }
        }
    }
}