package com.easy.convert.util

import com.aspose.words.*
import com.easy.convert.config.WordProperties
import com.easy.convert.model.ImageData
import com.easy.convert.model.TableData
import java.awt.Color
import java.io.File
import java.util.regex.Pattern
import javax.imageio.ImageIO

/**
 * @author  周泽
 * @date Create in 14:01 2020/7/29
 * @Description 模板替换
 */
class WordTemplateUtil(private val wordProperties: WordProperties) {

    /**
     * 模板渲染(包括文字，图片，表格)
     *
     * @param templatePath 模板路径
     * @param data 数据
     * @param outputPath 输出路径
     * @throws Exception
     */
    @Throws(Exception::class)
    fun render(templatePath: String, data: Map<String, Any?>, outputPath: String, watermarksEnable:Boolean = wordProperties.watermarks.enable) {
        // 加载文档
        val doc = Document(templatePath)
        // 按占位符前缀分割成多个段落
        val runs: Array<String> = doc.range.text.split(wordProperties.template.grammarPrefix).toTypedArray()
        // 循环渲染
        for (run in runs) {
            if (run.indexOf(wordProperties.template.grammarSuffix) != -1) {
                val name = run.substring(0, run.indexOf(wordProperties.template.grammarSuffix))
                when {
                    // 处理图片替换
                    name.startsWith(wordProperties.template.imageSymbol) -> renderImage(doc, data, name.substring(wordProperties.template.imageSymbol.length))
                    name.startsWith(wordProperties.template.tableSymbol) -> renderTable(doc, data, name.substring(wordProperties.template.imageSymbol.length))
                    name.startsWith(wordProperties.template.underLineSymbol) -> renderText(doc, data, name.substring(wordProperties.template.underLineSymbol.length), true)
                    else -> renderText(doc, data, name.substring(wordProperties.template.textSymbol.length), false)
                }
            } else {
                continue
            }
        }
        // 判断水印状态
        if (watermarksEnable) {
            WordConvertUtil(wordProperties).addConfigWatermarks(doc)
        }

        WordConvertUtil(wordProperties).checkSuffixAndSave(doc,outputPath)
    }

    /**
     * 重载,兼容java
     */
    fun render(templatePath: String, data: Map<String, Any?>, outputPath: String){
        render(templatePath, data, outputPath,wordProperties.watermarks.enable)
    }

    /**
     * 渲染文本
     * @param doc
     * @param data
     * @param name
     * @param underLine 是否添加下划线
     */
    private fun renderText(doc: Document, data: Map<String, Any?>, name: String, underLine: Boolean) {
        val value = data[name]
        // 拼接关键字
        var placeholder = wordProperties.template.prefixRegex + wordProperties.template.textSymbol + escapeExprSpecialWord(name) + wordProperties.template.suffixRegex
        if (value == null){
            // 内容是空,替换成空字符串
            doc.range.replace(Pattern.compile(placeholder),"")
            return
        }

        // 判断下划线
        if (underLine){
            val realName = wordProperties.template.underLineSymbol + name
            placeholder = wordProperties.template.prefixRegex + wordProperties.template.textSymbol + escapeExprSpecialWord(realName) + wordProperties.template.suffixRegex
        }

        val textData = { e: ReplacingArgs ->
            val builder = DocumentBuilder(e.matchNode.document as Document)
            builder.moveTo(e.matchNode)
            val font = builder.font
            // 设置文字颜色
//			font.setColor(Color.RED);
            //设置文字背景色
//            font.highlightColor = Color.yellow
            // 文字下划线
            if (underLine) {
                font.underline = 1
            }
            builder.write(value.toString())
            e.replacement = ""
            ReplaceAction.REPLACE
        }

        // 替换文本
        doc.range.replace(Pattern.compile(placeholder), textData, false)
    }

    /**
     * 渲染表格
     * @param doc
     * @param data
     * @param name
     */
    private fun renderTable(doc: Document, data: Map<String, Any?>, name: String) {
        val value = data[name]
        // 拼接关键字
        val placeholder = wordProperties.template.prefixRegex + "\\" + wordProperties.template.tableSymbol + escapeExprSpecialWord(name) + wordProperties.template.suffixRegex
        if (value == null || value !is TableData){
            // 内容是空,替换成空字符串
            doc.range.replace(Pattern.compile(placeholder),"")
            return
        }

        val tableData = { e:ReplacingArgs ->
            val builder = DocumentBuilder(e.matchNode.document as Document)
            builder.moveTo(e.matchNode)
            val table = builder.startTable()
            builder.bold = true
            table.setBorders(LineStyle.SINGLE, value.border, Color.BLACK)
            val headers: List<String> = value.headers
            for (header in headers) {
                builder.insertCell()
                builder.write(header)
                builder.cellFormat.verticalAlignment = CellVerticalAlignment.CENTER
                builder.paragraphFormat.alignment = ParagraphAlignment.CENTER
            }

            builder.endRow()
            builder.bold = false
            val rows: List<TableData.Row> = value.getRows()
            for (row in rows) {
                for (td in row.rowData) {
                    builder.insertCell()
                    builder.cellFormat.verticalAlignment = CellVerticalAlignment.CENTER
                    builder.paragraphFormat.alignment = ParagraphAlignment.LEFT
                    builder.write(td)
                }
                builder.endRow()
            }
            builder.endTable()
            e.replacement = ""
            ReplaceAction.REPLACE
        }

        // 替换表格
        doc.range.replace(Pattern.compile(placeholder), tableData, false)
    }

    /**
     * 渲染图片
     * @param doc
     * @param data
     * @param name
     */
    private fun renderImage(doc: Document, data: Map<String, Any?>, name: String) {
        val value = data[name]
        // 拼接关键字
        val placeholder = wordProperties.template.prefixRegex + "\\" + wordProperties.template.imageSymbol + escapeExprSpecialWord(name) + wordProperties.template.suffixRegex
        if (value == null || value !is com.easy.convert.model.ImageData){
            // 内容是空,替换成空字符串
            doc.range.replace(Pattern.compile(placeholder),"")
            return
        }

        // 声明图片的匿名内部类
        val imageData = { e: ReplacingArgs ->
            val builder = DocumentBuilder(e.matchNode.document as Document)
            builder.moveTo(e.matchNode)
            builder.insertImage(ImageIO.read(File(value.path)), value.width, value.height)
            e.replacement = ""
            ReplaceAction.REPLACE
        }
        // 替换
        doc.range.replace(Pattern.compile(placeholder),imageData,false)
    }

    /**
     * 提取模板中的全部关键字
     */
    fun findAllExprSpecialWord(filePath:String):String{
        val doc = Document(filePath)
        val sb = StringBuffer()
        // 获取所有段落
        val runs: Array<String> = doc.range.text.split(wordProperties.template.grammarPrefix).toTypedArray()
        for (run in runs) {
            if (run.indexOf(wordProperties.template.grammarSuffix) != -1) {
                val name = run.substring(0, run.indexOf(wordProperties.template.grammarSuffix))
                sb.append("${wordProperties.template.grammarPrefix}$name${wordProperties.template.grammarSuffix},")
            } else {
                continue
            }
        }
        return if (sb.isNotEmpty()) {
            sb.toString().substring(0, sb.toString().length - 1)
        } else sb.toString()
    }

}

/**
 * 替换特殊字符
 */
fun escapeExprSpecialWord(keyword: String): String {
    require(keyword.isNotEmpty()){ "关键字不能为空"}

    var convertKeyword = keyword
    val fbsArr = arrayOf("\\", "$", "(", ")", "*", "+", ".", "[", "]", "?", "^", "{", "}", "|")
    for (key in fbsArr) {
        if (convertKeyword.contains(key)) {
            convertKeyword = convertKeyword.replace(key, "\\" + key)
        }
    }
    return convertKeyword
}

fun main(args: Array<String>) {
    val mutableMapOf = mutableMapOf("name" to "zhangsan", "age" to 19)

    val imageData = ImageData("C:\\Users\\Administrator\\Desktop\\test.png", 200.00, 200.00)

    val headers = listOf("姓名", "年龄", "地址")

    val row1 = listOf("张三", "20", "杭州")
    val row2 = listOf("李四", "20", "杭州")
    val rows = listOf(TableData.Row(row1),TableData.Row(row2))
    val tableData = TableData(headers,rows)

    mutableMapOf["table1"] = tableData
    mutableMapOf["pic1"] = imageData

//    println(WordTemplateUtil(wordProperties = WordProperties()).findAllExprSpecialWord("C:\\Users\\Administrator\\Desktop\\xxxx1.docx"))

//    WordTemplateUtil(wordProperties = WordProperties()).render("C:\\Users\\Administrator\\Desktop\\xxxx1.docx",mutableMapOf,"C:\\Users\\Administrator\\Desktop\\xxxx1-out.docx")
}

