package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;

import javax.annotation.Nonnull;
import java.util.List;

public class TestListener implements ExcelReader.Listener {

    @Override
    public void onSheetStart(@Nonnull ExcelSheet sheet) {
    }

    @Override
    public void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        if (sheet.getShtIndex() == 0) {
            System.out.println("解析row：rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
            if (row.getRowIndex() < 7) {
                assert row.getColFirst() == 0;
                assert row.getColLast() == row.getRowIndex();
                assert cells.size() == row.getRowIndex() + 1;
                for (int i = 0; i < cells.size(); i++) {
                    ExcelCell cell = cells.get(i);
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    assert cell.getShtIndex() == sheet.getShtIndex();
                    assert cell.getRowIndex() == row.getRowIndex();
                    assert cell.getColIndex() == i;
                    assert cell.getStrValue().equals((char) ('A' + i) + "" + (row.getRowIndex() + 1));
                }
            } else {
                assert row.getColFirst() == row.getRowIndex() - 6;
                assert row.getColLast() == 6;
                assert cells.size() == 13 - row.getRowIndex();
                for (int i = 0; i < cells.size(); i++) {
                    ExcelCell cell = cells.get(i);
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    assert cell.getShtIndex() == sheet.getShtIndex();
                    assert cell.getRowIndex() == row.getRowIndex();
                    assert cell.getColIndex() == i + row.getColFirst();
                    assert cell.getStrValue().equals((char) ('A' + i + row.getColFirst()) + "" + (row.getRowIndex() + 1));
                }
            }
        } else {
            System.out.println("解析row：rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
            if (row.getRowIndex() < 3) {
                assert row.getColFirst() == 0;
                assert row.getColLast() == row.getRowIndex();
                assert cells.size() == row.getRowIndex() + 1;
                for (int i = 0; i < cells.size(); i++) {
                    ExcelCell cell = cells.get(i);
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    assert cell.getShtIndex() == sheet.getShtIndex();
                    assert cell.getRowIndex() == row.getRowIndex();
                    assert cell.getColIndex() == i;
                    assert cell.getStrValue().equals((row.getRowIndex() + 1) + "" + (char) ('A' + i));
                }
            } else {
                assert row.getColFirst() == row.getRowIndex() - 2;
                assert row.getColLast() == 2;
                assert cells.size() == 5 - row.getRowIndex();
                for (int i = 0; i < cells.size(); i++) {
                    ExcelCell cell = cells.get(i);
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    assert cell.getShtIndex() == sheet.getShtIndex();
                    assert cell.getRowIndex() == row.getRowIndex();
                    assert cell.getColIndex() == i + row.getColFirst();
                    assert cell.getStrValue().equals((row.getRowIndex() + 1) + "" + (char) ('A' + i + row.getColFirst()));
                }
            }
        }
    }

    @Override
    public void onSheetEnd(@Nonnull ExcelSheet sheet) {
    }
}