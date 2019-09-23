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
     * 流式导出excel 03或07版本，自动关闭输出流
     *
     * @param outStream excel文件输出流
     * @param provider  流式导出数据源
     * @throws Exception 导出中发生的异常
     */
    void write(@Nonnull OutputStream outStream, @Nonnull Provider provider) throws Exception;

    /**
     * 导出的excel版本
     */
    enum Version {
        /**
         * 03版本的Excel，虚拟流式导出
         */
        XLS,
        /**
         * 07版本的Excel，真实流式导出
         */
        XLSX
    }

    /**
     * 流式导出数据源
     */
    interface Provider {
        /**
         * 导出的excel版本
         *
         * @return excel版本
         */
        @Nonnull
        Version version();

        /**
         * 是否还有sheets信息
         *
         * @return sheets信息，返回null结束excel导出
         */
        @Nullable
        ExcelSheet provideSheet(int lastSheetIndex);

        /**
         * 是否还有row信息
         *
         * @param sheet        当前sheet
         * @param lastRowIndex 上一个row的序号，初始为-1
         * @return row信息，返回null则结束当前sheet写出
         */
        @Nullable
        ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex);

        /**
         * 提供某行的所有单元格，必须按顺序，可以不连续
         *
         * @param sheet 当前sheet
         * @param row   当前row
         * @return 某行的所有单元格
         */
        @Nonnull
        List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row);
    }
}