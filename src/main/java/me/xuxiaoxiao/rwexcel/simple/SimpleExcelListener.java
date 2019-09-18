package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/18 13:13]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleExcelListener<T> implements ExcelReader.Listener {
    private final int cache;
    private final List<T> list;

    private ExcelSheet sheet = null;
    private int rowStart = -1, rowEnd = -1;
    private List<ExcelCell> cells = new LinkedList<>();

    public SimpleExcelListener() {
        this(100);
    }

    public SimpleExcelListener(int cache) {
        this.cache = cache;
        this.list = new ArrayList<>(cache);
    }

    @Override
    public void onSheet(@Nonnull ExcelSheet sheet) {
        this.sheet = sheet;
    }

    @Override
    public void onRow(@Nonnull ExcelRow row) {
        int shtIndex = sheet.getShtIndex();
        while (rowEnd < row.getRowIndex() - 1) {
            //处理空白行
            if (!rowSkip(shtIndex, rowEnd, null)) {
                list.add(rowData(shtIndex, rowEnd, null, null));
                if (list.size() >= cache) {
                    flush();
                }
            }
            rowEnd++;
        }
        if (!rowSkip(shtIndex, row.getRowIndex(), row)) {
            list.add(rowData(shtIndex, row.getRowIndex(), row, cells));
            if (list.size() >= cache) {
                flush();
            }
        }
        cells.clear();
    }

    @Override
    public void onCell(@Nonnull ExcelCell cell) {
        cells.add(cell);
    }

    private void flush() {
        onList(sheet.getShtIndex(), rowStart, rowEnd, list);
        rowStart = rowEnd;
        list.clear();
    }

    protected boolean rowSkip(int shtIndex, int rowIndex, @Nullable ExcelRow row) {
        //第一行和空白行不处理，跳过
        return rowIndex == 0 || row == null;
    }

    protected T rowData(int shtIndex, int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) {
        return null;
    }

    protected abstract void onList(int shtIndex, int rowStart, int rowEnd, List<T> list);
}
