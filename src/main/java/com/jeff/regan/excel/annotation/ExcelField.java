//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

package com.jeff.regan.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.METHOD, ElementType.FIELD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {
    String value() default "";

    String title();

    int type() default 0;

    int align() default 0;

    int sort() default 0;

    String dictType() default "";

    Class<?> fieldType() default Class.class;

    int[] groups() default {};
}
