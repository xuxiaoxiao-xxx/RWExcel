package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.OutputStream;
import java.util.List;

/**
 * Excel导出器
 * <ul>
 * <li>[2019/9/12 18:22]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public interface ExcelWriter {
    /**
     * 流式写出excel 03或07版本
     *
     * @param outStream excel文件输出流
     * @param provider  流式写出数据源
     * @throws Exception 写出中发生的异常
     */
    void write(@Nonnull OutputStream outStream, @Nonnull Provider provider) throws Exception;

    enum Type {XLS, XLSX}

    /**
     * 流式写出数据源
     */
    interface Provider {
        /**
         * 写出的excel版本
         *
         * @return excel版本
         */
        @Nonnull
        Type version();

        /**
         * 所有sheets信息
         *
         * @return 所有sheets信息
         */
        @Nonnull
        ExcelSheet[] sheets();

        /**
         * 是否还有row数据
         *
         * @param sheet        当前sheet
         * @param lastRowIndex 上一个row的序号，初始为-1
         * @return 为null则结束当前sheet写出，不为null则继续写出新的row
         */
        @Nullable
        ExcelRow provideRow(ExcelSheet sheet, int lastRowIndex);

        /**
         * 为特定位置的cell提供数据
         *
         * @param sheet 当前sheet
         * @param row   当前row
         * @return 为null则写出空白单元格，不为null则继续写出数据
         */
        @Nonnull
        List<ExcelCell> provideCells(ExcelSheet sheet, ExcelRow row);
    }
}