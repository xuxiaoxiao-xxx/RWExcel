package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;

import javax.annotation.Nonnull;
import java.util.List;

/**
 * 简单Excel解析监听器
 * <ul>
 * <li>[2019/9/18 13:13]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleExcelListener implements ExcelReader.Listener {
    private volatile ExcelReader.Listener current;

    @Override
    public final void onSheetStart(@Nonnull ExcelSheet sheet) {
        this.current = sheetListener(sheet);
        if (this.current != null) {
            this.current.onSheetStart(sheet);
        }
    }

    @Override
    public final void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        if (this.current != null) {
            this.current.onRow(sheet, row, cells);
        }
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        if (this.current != null) {
            this.current.onSheetEnd(sheet);
        }
    }

    public abstract ExcelReader.Listener sheetListener(ExcelSheet sheet);
}