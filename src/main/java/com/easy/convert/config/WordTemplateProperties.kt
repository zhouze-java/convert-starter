package com.easy.convert.config

import com.easy.convert.util.escapeExprSpecialWord

/**
 * @author  周泽
 * @date Create in 14:02 2020/7/29
 * @Description
 */
data class WordTemplateProperties(
        // 语法前缀
        var grammarPrefix:String = "\${",
        // 语法后缀
        var grammarSuffix:String = "}",
        // 文本标识
        var textSymbol:String = "",
        // 图片标识
        var imageSymbol:String = "@",
        // 表格标识
        var tableSymbol:String = "#",
        // 下划线标识
        var underLineSymbol:String = "-"
){
    // 替换后的前缀
    val prefixRegex:String
        get() = escapeExprSpecialWord(grammarPrefix)

    // 替换后的后缀
    val suffixRegex:String
        get() = escapeExprSpecialWord(grammarSuffix)

}
