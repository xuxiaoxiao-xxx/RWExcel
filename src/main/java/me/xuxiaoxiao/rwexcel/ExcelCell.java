package me.xuxiaoxiao.rwexcel;

import lombok.Data;

import javax.annotation.Nonnull;
import java.io.Serializable;

/**
 * 通用Cell信息
 * <ul>
 * <li>[2019/9/12 19:42]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
@Data
public class ExcelCell implements Serializable {
    /**
     * Sheet编号，从0开始
     */
    private final int shtIndex;
    /**
     * Row编号，从0开始
     */
    private final int rowIndex;
    /**
     * Cell编号，从0开始
     */
    private final int colIndex;
    /**
     * Cell的字符串值，不为null
     */
    @Nonnull
    private String strValue;
}
