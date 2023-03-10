package com.easy.convert.util

/**
 * @author  周泽
 * @date Create in 16:32 2020/7/29
 * @Description 公共方法
 */

/**
 * 获取文件后缀名
 */
fun getFileSuffix(filePath: String):String{
    return filePath.substring(filePath.lastIndexOf(".") + 1)
}