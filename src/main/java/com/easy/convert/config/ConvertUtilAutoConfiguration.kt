package com.easy.convert.config

import com.easy.convert.util.excel.ExcelConvertUtil
import com.easy.convert.util.word.WordConvertUtil
import com.easy.convert.util.word.WordTemplateUtil
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.boot.context.properties.EnableConfigurationProperties
import org.springframework.context.annotation.Bean
import org.springframework.context.annotation.Configuration

/**
 * @author  周泽
 * @date Create in 11:25 2020/7/28
 * @Description 自动配置类
 */
@Configuration
@EnableConfigurationProperties(WordProperties::class)
open class ConvertUtilAutoConfiguration{

    @Autowired
    lateinit var wordProperties: WordProperties

    @Bean
    open fun wordConvertUtil(): WordConvertUtil {
        return WordConvertUtil(wordProperties)
    }

    @Bean
    open fun wordTemplateUtil(): WordTemplateUtil {
        return WordTemplateUtil(wordProperties)
    }

    @Bean
    open fun excelConvertUtil():ExcelConvertUtil{
        return ExcelConvertUtil()
    }

}