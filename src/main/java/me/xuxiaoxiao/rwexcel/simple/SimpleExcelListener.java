package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.ArrayList;
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
    private final ArrayList<SimpleSheetListener<?>> listeners = new ArrayList<>();

    private ExcelSheet sheet = null;

    @Override
    public final void onSheetStart(@Nonnull ExcelSheet sheet) {
        this.sheet = sheet;
        while (this.sheet.getShtIndex() >= this.listeners.size()) {
            this.listeners.add(sheetListener(this.sheet));
        }
        SimpleSheetListener<?> sheetListener = this.listeners.get(this.sheet.getShtIndex());
        if (sheetListener != null) {
            sheetListener.onSheetStart(this.sheet);
        }
    }

    @Override
    public final void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        SimpleSheetListener<?> sheetListener = this.listeners.get(this.sheet.getShtIndex());
        if (sheetListener != null) {
            sheetListener.onRow(sheet, row, cells);
        }
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        SimpleSheetListener<?> sheetListener = this.listeners.get(this.sheet.getShtIndex());
        if (sheetListener != null) {
            sheetListener.onSheetEnd(this.sheet);
        }
    }

    /**
     * 获取对应sheet的监听器
     *
     * @param sheet sheet信息
     * @return 对应sheet的监听器
     */
    @Nullable
    public abstract SimpleSheetListener<?> sheetListener(@Nonnull ExcelSheet sheet);
}