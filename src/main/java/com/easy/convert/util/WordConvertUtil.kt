package com.easy.convert.util

import com.aspose.words.*
import com.easy.convert.config.WordProperties
import com.easy.convert.config.WordWatermarksProperties
import java.util.*

/**
 * @author  周泽
 * @date Create in 10:48 2020/7/28
 * @Description word格式转换
 */
class WordConvertUtil(private val wordProperties: WordProperties) {

    /**
     * @param filePath 文件路径,
     * @param saveFormat 转换后的格式, 支持 doc,docx,pdf,html,text,jpeg, 不区分大小写
     * @param outputFilePath 输出路径,除图片外具体到文件名称,图片具体到文件夹即可
     * @param resolution 转换图片时的质量
     */
    @Throws(Exception::class)
    fun convertFormat(filePath: String, saveFormat: String, outputFilePath: String, resolution:Float = 100f): MutableList<String> {
        require(filePath.endsWith(".doc") || filePath.endsWith(".docx")) {"输入文件必须是word, 后缀为 doc 或 docx"}
        val format: Int
        when {
            saveFormat.equals("doc", ignoreCase = true) -> {
                format = SaveFormat.DOC
            }
            saveFormat.equals("docx", ignoreCase = true) -> {
                format = SaveFormat.DOCX
            }
            saveFormat.equals("pdf", ignoreCase = true) -> {
                format = SaveFormat.PDF
            }
            saveFormat.equals("html", ignoreCase = true) -> {
                format = SaveFormat.HTML
            }
            saveFormat.equals("text", ignoreCase = true) -> {
                format = SaveFormat.TEXT
            }
            saveFormat.equals("jpeg", ignoreCase = true) -> {
                return convertImg(filePath, outputFilePath, resolution = resolution)
            }
            saveFormat.equals("png",ignoreCase = true) -> {
                return convertImg(filePath, outputFilePath, SaveFormat.PNG,resolution)
            }
            else -> {
                throw IllegalArgumentException("不支持的类型:$saveFormat")
            }
        }

        // 加载源文件
        val doc = Document(filePath)
        // 判断是否需要添加水印
        if (wordProperties.watermarks.enable){
            addConfigWatermarks(doc)
        }
        // 保存
        doc.save(outputFilePath, format)
        // 之后返回路径
        return mutableListOf(outputFilePath)
    }

    /**
     * 重载,兼容Java
     */
    fun convertFormat(filePath: String, saveFormat: String, outputFilePath: String): MutableList<String> {
        return convertFormat(filePath, saveFormat, outputFilePath,100f)
    }

    /**
     * @param filePath 文件路径
     * @param outputFilePath 输出路径,只具体到文件夹即可
     * @param format 图片格式,默认jpeg
     * @param resolution 图片质量,默认100f
     * @return 返回全部路径
     */
    private fun convertImg(filePath: String, outputFilePath: String, format: Int = SaveFormat.JPEG, resolution: Float = 100f): MutableList<String> {
        // 加载文档
        val doc = Document(filePath)
        // 是否需要添加水印
        if (wordProperties.watermarks.enable) {
            addConfigWatermarks(doc)
        }
        // 基础信息设置
        val iso = ImageSaveOptions(format)
        iso.resolution = resolution
        iso.prettyFormat = true
        iso.useAntiAliasing = true
        // 存放图片路径
        val imagesPath: MutableList<String> = ArrayList()
        val fileName = "image_" + UUID.randomUUID()

        // 循环转换图片
        for (i in 0 until doc.pageCount) {
            iso.pageIndex = i
            doc.save("$outputFilePath$fileName$i.jpeg", iso)
            imagesPath.add("$fileName$i.jpeg")
        }

        return imagesPath
    }

    /**
     * 将多个 word 追加到第一个word中去
     * @param wordPath 要合并的word路径集合
     * @param outPath 输出路径,具体到文件名
     */
    fun appendDoc(wordPath: List<String>, outPath: String) {
        require(wordPath.isNotEmpty()) { "输入word路径集合不能为空" }

        // 把所有的doc找出来放到一个集合中来
        val wordList = mutableListOf<Document>()
        for (path in wordPath) {
            val document = Document(path)
            wordList.add(document)
        }

        val firstDoc = wordList[0]
        for (document in wordList) {
            firstDoc.appendDocument(document, ImportFormatMode.USE_DESTINATION_STYLES)
        }

        // 检查后缀并保存
        checkSuffixAndSave(firstDoc,outPath)
    }

    /**
     * 为文档添加水印,水印信息取自Spring的配置
     */
    fun addConfigWatermarks(document: Document, watermarksProperties:WordWatermarksProperties = this.wordProperties.watermarks){
        // 创建一个水印
        val watermark = Shape(document, ShapeType.TEXT_PLAIN_TEXT)

        // 设置水印的相关属性, 文本,字体,宽高,倾斜,字体颜色
        watermark.textPath.text = watermarksProperties.text
        watermark.textPath.fontFamily = watermarksProperties.fontFamily
        watermark.width = watermarksProperties.width
        watermark.height = watermarksProperties.height
        watermark.rotation = watermarksProperties.rotation
        watermark.fill.color = watermarksProperties.color
        watermark.strokeColor = watermarksProperties.color

        // 设置相对水平位置,垂直位置,包装类型,垂直对齐,水平对齐,默认写死,可以写到配置类中实现自定义
        watermark.relativeHorizontalPosition = RelativeHorizontalPosition.PAGE
        watermark.relativeVerticalPosition = RelativeVerticalPosition.PAGE
        watermark.wrapType = WrapType.NONE
        watermark.verticalAlignment = VerticalAlignment.CENTER
        watermark.horizontalAlignment = HorizontalAlignment.CENTER

        // 创建一个段落,然后把水印添加到新的段落中去
        val watermarkPara = Paragraph(document)
        watermarkPara.appendChild(watermark)

        // 循环添加到文档的每个位置
        for (sect in document.sections) {
            // the watermark to appear on all pages, insert into all headers.
            insertWatermarkIntoHeader(watermarkPara,sect,HeaderFooterType.HEADER_PRIMARY)
        }
    }

    private fun insertWatermarkIntoHeader(watermarkPara: Paragraph, sect: Section, headerType: Int) {
        var header = sect.headersFooters.getByHeaderFooterType(headerType)

        if (header == null) {
            // 当前部分中没有指定类型的标题,创建
            header = HeaderFooter(sect.document, headerType)
            sect.headersFooters.add(header)
        }
        // 插入水印的克隆到头部

        // Insert a clone of the watermark into the header.
        header.appendChild(watermarkPara.deepClone(true))
    }

    /**
     * 添加自定义水印
     * @param filePath 文档路径
     * @param watermarksProperties 水印详细配置
     */
    fun addCustomWatermarks(filePath: String, outputFilePath: String, watermarksProperties: WordWatermarksProperties){
        require(filePath.endsWith(".doc")||filePath.endsWith(".docx"))
        val document = Document(filePath)
        // 添加水印
        addConfigWatermarks(document,watermarksProperties)
        // 检查后缀并保存
        checkSuffixAndSave(document, outputFilePath)
    }

    /**
     * 检查文件后缀名,并保存
     * @param document 文档
     * @param outputFilePath 输出路径
     */
    fun checkSuffixAndSave(document: Document, outputFilePath: String) {
        // 拿到文件后缀名
        val suffix = getFileSuffix(outputFilePath)
        // 判断保存
        if (suffix == "doc") {
            document.save(outputFilePath, SaveFormat.DOC)
        }
        if (suffix == "docx") {
            document.save(outputFilePath, SaveFormat.DOCX)
        }
    }

    /**
     * 克隆文件
     * @param filePath 输入路径
     * @param outputFilePath 输出路径
     * @param watermarks 是否添加水印,默认是全局配置
     */
    fun clone(filePath: String,outputFilePath: String, enableWatermarks:Boolean = wordProperties.watermarks.enable){
        val document = Document(filePath)
        // 克隆
        val cloneDocument = document.deepClone()
        // 判断水印状态
        if (enableWatermarks) {
            addConfigWatermarks(cloneDocument)
        }
        // 检查后缀并保存
        checkSuffixAndSave(cloneDocument, outputFilePath)
    }

    /**
     * 重载,兼容Java
     */
    fun clone(filePath: String,outputFilePath: String){
        clone(filePath, outputFilePath,wordProperties.watermarks.enable)
    }

    /**
     * 不加水印复制
     */
    fun cloneWithNoWatermarks(filePath: String,outputFilePath: String){
        clone(filePath, outputFilePath, false)
    }

    /**
     * txt 转 word
     * @param filePath
     * @param outputFilePath
     */
    fun txtToWord(filePath: String,outputFilePath: String){
        require(getFileSuffix(filePath) == "txt"){"输入文件必须是txt"}

        // 直接把txt加载成document
        val document = Document(filePath)

        // 保存
        document.save(outputFilePath)
    }

    /**
     * 获取文件后缀名
     */
    private fun getFileSuffix(filePath: String):String{
        return filePath.substring(filePath.lastIndexOf(".") + 1)
    }

    /**
     * 修改文档为只读
     * @param filePath 输入路径
     * @param outputFilePath 输出路径
     */
    fun toReadOnly(filePath: String, outputFilePath: String) {
        protect(filePath,outputFilePath,ProtectionType.READ_ONLY)
    }

    /**
     * 文档保护
     * @param filePath 输入路径
     * @param outputFilePath 输出路径
     * @param protectionType 保护等级,具体参考 {@link ProtectionType} 类
     */
    fun protect(filePath: String, outputFilePath: String, protectionType:Int){
        val document = Document(filePath)
        // 修改
        document.protect(protectionType)
        // 输出
        checkSuffixAndSave(document, outputFilePath)
    }


}

fun main(args: Array<String>) {
    val wordWatermarksProperties = WordWatermarksProperties()
    wordWatermarksProperties.text = "test"

//    WordConvertUtil(WordProperties()).clone("C:\\Users\\Administrator\\Desktop\\xxxx.docx", "C:\\Users\\Administrator\\Desktop\\xxxx-1.docx")
    WordConvertUtil(WordProperties()).toReadOnly("C:\\Users\\Administrator\\Desktop\\xxxx.docx","C:\\Users\\Administrator\\Desktop\\test-1.docx")
}

