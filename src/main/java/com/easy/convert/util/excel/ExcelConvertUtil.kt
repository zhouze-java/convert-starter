package com.easy.convert.util.excel

import com.aspose.cells.*
import java.io.File
import java.io.InputStream
import java.io.OutputStream


/**
 * @author  周泽
 * @date Create in 11:18 2023/3/9
 * @Description
 */
class ExcelConvertUtil {

    /**
     * 表格太宽的话会丢失, 如果是列是固定的, 宽度不会太大的话可以尝试使用
     */
    @Throws(Exception::class)
    fun convertFirstSheetToPng(excelInputStream: InputStream, outputStream: OutputStream) {
        if (!authLicense()) {
            throw Exception("excel license is not available")
        }

        val workbook = Workbook(excelInputStream)
        val worksheet = workbook.worksheets[0]

        settingSheetCellAdaptiveAndBorder(worksheet, workbook)


        val imageOrPrintOptions = ImageOrPrintOptions()

        imageOrPrintOptions.imageType = ImageType.PNG
        imageOrPrintOptions.isCellAutoFit = true
        imageOrPrintOptions.horizontalResolution = 300
        imageOrPrintOptions.verticalResolution = 300
        imageOrPrintOptions.textCrossType = TextCrossType.CROSS_KEEP
        imageOrPrintOptions.gridlineType  = GridlineType.DOTTED



        val sheetRender = SheetRender(worksheet, imageOrPrintOptions)

        sheetRender.toImage(0, outputStream)
    }



    /**
     * 转html
     */
    fun convertToHtml(excelInputStream: InputStream, outputStream: OutputStream) {
        if (!authLicense()) {
            throw Exception("excel license is not available")
        }


        val workbook = Workbook(excelInputStream)

        settingSheetCellAdaptiveAndBorder(workbook.worksheets[0], workbook)

        val htmlSaveOptions = HtmlSaveOptions()
        htmlSaveOptions.exportGridLines = true
        workbook.save(outputStream, htmlSaveOptions)
    }

    /**
     * 调用方法之前都要调用这个, 不然会有水印
     */
    private fun authLicense(): Boolean {
        var result = false
        try {
            val `is` = License::class.java.getResourceAsStream("/com.aspose.cells.lic_2999.xml")
            val asposeLicense = License()
            asposeLicense.setLicense(`is`)
            `is`?.close()
            result = true
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return result
    }

    private fun settingSheetCellAdaptiveAndBorder(worksheet: Worksheet, workbook: Workbook) {

        // 设置表格宽高自适应, 否则图片会显示不全
        val cells = worksheet.cells
        val maxDisplayRange = cells.maxDisplayRange
        val autoFitterOptions = AutoFitterOptions()
        autoFitterOptions.autoFitMergedCellsType = AutoFitMergedCellsType.EACH_LINE
        autoFitterOptions.autoFitWrappedTextType = AutoFitWrappedTextType.PARAGRAPH

        worksheet.autoFitColumns(maxDisplayRange.firstColumn, maxDisplayRange.columnCount, autoFitterOptions)
        worksheet.autoFitRows(maxDisplayRange.firstRow, maxDisplayRange.rowCount, autoFitterOptions)

        // 设置单元格的边框, 和自动换行
        val style = workbook.createStyle()
        style.setBorder(BorderType.TOP_BORDER, CellBorderType.THIN, Color.getBlack())
        style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.THIN, Color.getBlack())
        style.setBorder(BorderType.LEFT_BORDER, CellBorderType.THIN, Color.getBlack())
        style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.THIN, Color.getBlack())
        style.isTextWrapped = true

        val styleFlag = StyleFlag()
        styleFlag.all = true

        cells.applyStyle(style, styleFlag)
    }
}

fun main() {
    ExcelConvertUtil().convertToHtml(
        File("C:\\Users\\周泽\\Desktop\\test1.xlsx").inputStream(),
        File("C:\\Users\\周泽\\Desktop\\test.html").outputStream()
    )
}