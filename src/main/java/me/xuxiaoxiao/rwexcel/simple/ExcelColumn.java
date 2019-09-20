package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.simple.converter.Converter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * Excel列注解
 * <ul>
 * <li>[2019/9/19 8:48]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /**
     * 对应excel中的列序号，从0开始。
     * 默认为-1，当所有注解列序号都是-1时，将开启自动寻列模式，根据title和sheet第一行单元格内的文字自动匹配列和成员域。
     *
     * @return 对应excel中的列序号
     */
    int index() default -1;

    /**
     * 列标题，尽量保证唯一性，有利于自动寻列
     *
     * @return 列标题
     */
    String value();

    /**
     * 值转换器class对象
     *
     * @return 值转换器class对象
     */
    Class<? extends Converter> converter() default Converter.class;
}
