package com.easy.convert.annotation;

import java.lang.annotation.*;

/**
 * @author 周泽
 * @date Create in 15:01 2020/11/2
 * @Description 导出Excel
 */
@Target({ElementType.METHOD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ResponseExcel {

    String fileName() default "";

    /**
     * 文件后缀名
     * @return
     */
    String fileSuffix() default ".xlsx";

    /**
     * 模板url
     * @return
     */
    String templateUrl();

}
