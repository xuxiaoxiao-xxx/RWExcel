package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleExcelProvider implements ExcelWriter.Provider {
    private final ArrayList<SimpleSheetProvider<?>> providers = new ArrayList<>();

    @Nonnull
    @Override
    public ExcelWriter.Type version() {
        return ExcelWriter.Type.XLSX;
    }

    @Nonnull
    @Override
    public ExcelSheet[] sheets() {
        return new ExcelSheet[1];
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(ExcelSheet sheet, int lastRowIndex) {
        while (sheet.getShtIndex() >= this.providers.size()) {
            this.providers.add(sheetProvider(sheet));
        }
        SimpleSheetProvider<?> sheetProvider = this.providers.get(sheet.getShtIndex());
        if (sheetProvider != null) {
            return sheetProvider.provideRow(sheet, lastRowIndex);
        } else {
            return null;
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(ExcelSheet sheet, ExcelRow row) {
        SimpleSheetProvider<?> sheetProvider = this.providers.get(sheet.getShtIndex());
        if (sheetProvider != null) {
            return sheetProvider.provideCells(sheet, row);
        } else {
            return Collections.emptyList();
        }
    }

    public abstract SimpleSheetProvider<?> sheetProvider(ExcelSheet sheet);
}
