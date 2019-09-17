package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;

import javax.annotation.Nonnull;
import java.io.InputStream;

/**
 * Excel解析类
 * <ul>
 * <li>[2019/9/12 18:22]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public interface ExcelReader {

    /**
     * 流式读取excel 03或07版本
     *
     * @param inStream excel文件输入流
     * @param listener 流式读取监听器
     * @throws Exception 读取中发生的异常
     */
    void read(@Nonnull InputStream inStream, @Nonnull Listener listener) throws Exception;

    /**
     * 流式读取监听器
     */
    interface Listener {
        /**
         * 处理sheet
         *
         * @param sheet sheet信息
         */
        void onSheet(@Nonnull ExcelSheet sheet);

        /**
         * 处理行
         * <ul>
         *     <li>保证按顺序处理行，行号小的先处理</li>
         *     <li>不保证连续，遇到空行会跳过</li>
         *     <li>行内保证至少有一个单元格</li>
         * </ul>
         *
         * @param row row信息
         */
        void onRow(@Nonnull ExcelRow row);

        /**
         * 处理单元格
         * <ul>
         *     <li>保证按顺序处理行，列号小的先处理</li>
         *     <li>不保证连续，遇到空单元格（BlankRecord）会跳过</li>
         *     <li>所在行内保证至少有一个单元格</li>
         *     <li>保证读取单元格的值不为null</li>
         * </ul>
         *
         * @param cell cell信息
         */
        void onCell(@Nonnull ExcelCell cell);
    }
}
