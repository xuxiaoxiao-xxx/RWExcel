package me.xuxiaoxiao.rwexcel;

import lombok.Data;

import java.io.Serializable;

/**
 * 通用Row信息
 * <ul>
 * <li>[2019/9/12 19:42]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
@Data
public class ExcelRow implements Serializable {
    /**
     * Sheet编号，从0开始
     */
    private final int shtIndex;
    /**
     * Row编号，从0开始
     */
    private final int rowIndex;
    /**
     * Row第一个Cell的编号，从0开始
     */
    private int colFirst;
    /**
     * Row最后一个Cell的编号，从0开始
     */
    private int colLast;
}
