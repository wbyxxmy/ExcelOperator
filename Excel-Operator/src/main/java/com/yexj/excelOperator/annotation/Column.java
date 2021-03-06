package com.yexj.excelOperator.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Column {
    /**
     * 每一列都有一个名字
     */
    String name();

    /**
     * 导出列1,导入列-1,既导入又导出0
     */
     int export() default 0;

    /**
     * 排在前面的列具有低优先级，优先级相同时，按照出现的顺序排列
     */
    int priority() default 0;
}
