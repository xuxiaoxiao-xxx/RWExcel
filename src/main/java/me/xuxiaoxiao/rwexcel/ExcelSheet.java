package me.xuxiaoxiao.rwexcel;

import lombok.Data;

import javax.annotation.Nonnull;
import java.io.Serializable;

/**
 * 通用Sheet信息
 * <ul>
 * <li>[2019/9/12 19:42]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
@Data
public class ExcelSheet implements Serializable {
    /**
     * Sheet编号，从0开始
     */
    private final int shtIndex;
    /**
     * Sheet名称
     */
    @Nonnull
    private final String shtName;
}
