package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;

import javax.annotation.Nonnull;
import java.io.InputStream;
import java.util.List;

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
     * 流式读取excel 03或07版本，自动关闭输入流
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
         * 处理sheet开头
         *
         * @param sheet sheet信息
         */
        void onSheetStart(@Nonnull ExcelSheet sheet);

        /**
         * 处理行
         * <ul>
         *     <li>保证按顺序处理行，行号小的先处理，列号小的先处理</li>
         *     <li>不保证连续，遇到空行会跳过，遇到空单元格（BlankRecord）会跳过</li>
         *     <li>行内保证至少有一个单元格</li>
         *     <li>保证读取单元格的值不为null</li>
         * </ul>
         *
         * @param sheet sheet信息
         * @param row   row信息
         * @param cells cells信息
         */
        void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells);

        /**
         * 处理sheet结尾
         *
         * @param sheet sheet信息
         */
        void onSheetEnd(@Nonnull ExcelSheet sheet);
    }
}
