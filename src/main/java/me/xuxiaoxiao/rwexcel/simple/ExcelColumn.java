package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import me.xuxiaoxiao.rwexcel.simple.converter.StrConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/19 8:48]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    int index() default -1;

    String title();

    Class<? extends Converter<?>> converter() default StrConverter.class;
}
