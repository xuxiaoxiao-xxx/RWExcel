package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.Collections;
import java.util.List;

/**
 * 简单Excel导出数据源
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleExcelProvider implements ExcelWriter.Provider {
    private volatile SimpleSheetProvider<?> current;

    @Nonnull
    @Override
    public ExcelWriter.Version version() {
        return ExcelWriter.Version.XLSX;
    }

    @Nullable
    @Override
    public final ExcelSheet provideSheet(int lastSheetIndex) {
        this.current = sheetProvider(lastSheetIndex);
        if (this.current != null) {
            return this.current.provideSheet(lastSheetIndex);
        } else {
            return null;
        }
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
        if (this.current != null) {
            return current.provideRow(sheet, lastRowIndex);
        } else {
            return null;
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
        if (this.current != null) {
            return this.current.provideCells(sheet, row);
        } else {
            return Collections.emptyList();
        }
    }

    public abstract SimpleSheetProvider<?> sheetProvider(int lastSheetIndex);
}
