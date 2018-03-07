package com.yexj.excelOperator;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Sheet {
    /**
     * 每一张表需要一个表名
     */
    String name() default "sheet";

    /**
     * 是否输出表头
     */
    boolean containsHeader() default true;

    /**
     * 是否在第一列输出序号
     */
    boolean containsSequence() default false;
}
