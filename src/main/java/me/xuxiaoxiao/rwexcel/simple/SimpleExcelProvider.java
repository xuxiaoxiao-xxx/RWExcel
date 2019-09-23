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
 * 简单Excel导出数据源
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
    public ExcelWriter.Version version() {
        return ExcelWriter.Version.XLSX;
    }

    @Nullable
    @Override
    public ExcelSheet provideSheet(int lastSheetIndex) {
        if (lastSheetIndex < 0) {
            return new ExcelSheet(0, "Sheet1");
        } else {
            return null;
        }
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
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
    public final List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
        SimpleSheetProvider<?> sheetProvider = this.providers.get(sheet.getShtIndex());
        if (sheetProvider != null) {
            return sheetProvider.provideCells(sheet, row);
        } else {
            return Collections.emptyList();
        }
    }

    /**
     * 获取对应sheet的数据源
     *
     * @param sheet sheet信息
     * @return 对应sheet的数据源
     */
    @Nullable
    public abstract SimpleSheetProvider<?> sheetProvider(ExcelSheet sheet);
}
