package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.LinkedList;
import java.util.List;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetProvider<T> implements ExcelWriter.Provider {
    private final ExcelSheet sheet;
    private final List<T> list = new LinkedList<>();

    public SimpleSheetProvider(ExcelSheet sheet) {
        this.sheet = sheet;
    }

    @Nonnull
    @Override
    public ExcelWriter.Type version() {
        return ExcelWriter.Type.XLSX;
    }

    @Nonnull
    @Override
    public ExcelSheet[] sheets() {
        return new ExcelSheet[0];
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(ExcelSheet sheet, int lastRowIndex) {
        if (list.isEmpty()) {
            list.addAll(queryList(this.sheet, lastRowIndex));
        }
        if (list.isEmpty()) {
            return null;
        } else {
            return entityRow(sheet, lastRowIndex, list.get(0));
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(ExcelSheet sheet, ExcelRow row) {
        return entityCells(sheet, row, list.remove(0));
    }

    protected ExcelRow entityRow(ExcelSheet sheet, int lastRowIndex, T entity) {
        return new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
    }

    protected List<ExcelCell> entityCells(ExcelSheet sheet, ExcelRow row, T entity) {
        return null;
    }

    public abstract List<T> queryList(ExcelSheet sheet, int lastRowIndex);
}
