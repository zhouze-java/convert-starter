package com.easy.convert.config

import java.awt.Color


/**
 * @author  周泽
 * @date Create in 15:19 2020/7/28
 * @Description
 */
data class WordWatermarksProperties(
        // 启用水印
        var enable: Boolean = false,
        // 水印内容
        var text: String = "default-watermarks",
        // 字体设置
        var fontFamily: String = "Arial",
        // 水印宽
        var width: Double = 500.00,
        // 水印高
        var height: Double = 100.00,
        // 倾斜, 左下角到右上角
        var rotation:Double = -40.00,
        // 水印默认灰色
        var color:Color = Color.GRAY
        // ... 其他配置需要再加
)