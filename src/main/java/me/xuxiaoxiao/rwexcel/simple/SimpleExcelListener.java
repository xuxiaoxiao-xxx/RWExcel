package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;

import javax.annotation.Nonnull;
import java.util.ArrayList;
import java.util.List;

/**
 * 请填写类的描述
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
    public final void onRow(@Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        SimpleSheetListener<?> sheetListener = this.listeners.get(this.sheet.getShtIndex());
        if (sheetListener != null) {
            sheetListener.onRow(row, cells);
        }
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        SimpleSheetListener<?> sheetListener = this.listeners.get(this.sheet.getShtIndex());
        if (sheetListener != null) {
            sheetListener.onSheetEnd(this.sheet);
        }
    }

    public abstract SimpleSheetListener<?> sheetListener(ExcelSheet sheet);
}