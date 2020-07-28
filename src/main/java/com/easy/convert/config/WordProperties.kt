package com.easy.convert.config

import org.springframework.boot.context.properties.ConfigurationProperties
import org.springframework.stereotype.Component

/**
 * @author  周泽
 * @date Create in 14:59 2020/7/28
 * @Description
 */
@Component
@ConfigurationProperties(prefix = "file.config.word")
data class WordProperties(
        // 水印相关配置
        val watermarks: WordWatermarksProperties = WordWatermarksProperties()
)