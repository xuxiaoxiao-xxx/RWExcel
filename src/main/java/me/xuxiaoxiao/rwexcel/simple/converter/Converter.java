package me.xuxiaoxiao.rwexcel.simple.converter;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/19 17:29]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public interface Converter<T> {

    T convert(String str);
}
