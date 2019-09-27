package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.junit.Test;

import javax.annotation.Nonnull;
import java.util.List;
import java.util.Objects;

public class ExcelStreamReaderTest {

    @Test
    public void demo() throws Exception {
        new ExcelStreamReader().read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xls")), new ExcelReader.Listener() {
            @Override
            public void onSheetStart(@Nonnull ExcelSheet sheet) {
                System.out.println("开始解析Sheet：" + sheet.getShtName());
            }

            @Override
            public void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
                System.out.println("解析row：rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
                for (ExcelCell cell : cells) {
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                }
            }

            @Override
            public void onSheetEnd(@Nonnull ExcelSheet sheet) {
                System.out.println("结束解析Sheet：" + sheet.getShtName());
            }
        });
    }

    @Test
    public void readXls() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        TestListener listener = new TestListener();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xls")), listener);
        assert listener.sheets == 2;
    }

    @Test
    public void readXlsx() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        TestListener listener = new TestListener();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xlsx")), listener);
        assert listener.sheets == 2;
    }

    public static final class TestListener implements ExcelReader.Listener {
        public int sheets = 0;

        @Override
        public void onSheetStart(@Nonnull ExcelSheet sheet) {
            sheets++;
            System.out.println("解析Sheet" + sheets);
            assert sheet.getShtIndex() == sheets - 1;
            assert sheet.getShtName().equals("Sheet" + sheets);
        }

        @Override
        public void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
            assert sheet.getShtIndex() == sheets - 1;
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
            System.out.println("结束sheet：" + sheet.getShtName());
            System.out.println();
            assert sheet.getShtIndex() == sheets - 1;
            assert sheet.getShtName().equals("Sheet" + sheets);
        }
    }
}