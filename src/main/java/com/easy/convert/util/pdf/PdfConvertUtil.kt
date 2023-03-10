package com.easy.convert.util.pdf

import com.aspose.pdf.Document
import com.aspose.pdf.SaveFormat
import com.easy.convert.util.getFileSuffix

/**
 * @author  周泽
 * @date Create in 23:02 2020/7/29
 * @Description
 */
class PdfConvertUtil {
    fun convertFormat(filePath: String, outputFilePath: String) {
        require(getFileSuffix(filePath) == "pdf") { "输入文件必须是pdf格式" }

        val document = Document(filePath)
        document.save(outputFilePath, SaveFormat.DocX)

    }
}

fun main(args: Array<String>) {
    PdfConvertUtil().convertFormat("C:\\Users\\Administrator\\Desktop\\xxxx.pdf","C:\\Users\\Administrator\\Desktop\\xxxx-out.docx")
}