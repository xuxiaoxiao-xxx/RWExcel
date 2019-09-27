package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Objects;

public class SimpleExcelListenerTest {

    @Test
    public void demo() throws Exception {
        new ExcelStreamReader().read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheets.xls")), new SimpleExcelListener() {

            @Nullable
            @Override
            public SimpleSheetListener<TestEntity> sheetListener(@Nonnull ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    return new SimpleSheetListener<TestEntity>(sheet) {
                        @Override
                        protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                            System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                            for (TestEntity entity : list) {
                                System.out.println(entity);
                            }
                        }
                    };
                } else {
                    return null;
                }
            }
        });
    }

    @Test
    public void testSimpleSheets() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheets.xls")), new TestSheetsListener());
    }

    public static final class TestSheetsListener extends SimpleExcelListener {

        @Override
        public SimpleSheetListener<?> sheetListener(@Nonnull ExcelSheet sheet) {
            if (sheet.getShtIndex() == 0) {
                return new SimpleSheetListener<TestEntity>(sheet) {
                    @Override
                    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                        System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                        assert rowStart == 1;
                        assert rowEnd == 2;
                        assert list.size() == 1;
                        TestEntity entity = list.get(0);
                        System.out.println("解析到实体" + entity);
                        assert entity.getColStr().equals("str");
                        assert entity.getColInt() == 1;
                        assert entity.getColDbl() == 2;
                        assert entity.getColLng() == 3;
                        assert entity.getColFlt() == 4;
                        assert entity.isColBol();
                        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity.getColDat()).equals("2019-01-01 00:00:00.000");
                    }
                };
            } else {
                return new SimpleSheetListener<TestEntity>(sheet) {
                    @Override
                    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                        System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                        assert rowStart == 1;
                        assert rowEnd == 3;
                        assert list.size() == 1;
                        TestEntity entity = list.get(0);
                        System.out.println("解析到实体" + entity);
                        assert entity.getColStr().equals("str1");
                        assert entity.getColInt() == 2;
                        assert entity.getColDbl() == 3;
                        assert entity.getColLng() == 4;
                        assert entity.getColFlt() == 5;
                        assert !entity.isColBol();
                        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity.getColDat()).equals("2020-01-01 00:00:00.000");
                    }
                };
            }
        }
    }

    @Test
    public void testSimpleTitle() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleTitle.xls")), new SimpleExcelListener() {

            @Override
            public SimpleSheetListener<?> sheetListener(@Nonnull ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    return new SimpleSheetListener<TestEntity>(sheet) {
                        @Override
                        protected int titleRowCount() {
                            return 0;
                        }

                        @Nullable
                        @Override
                        protected TestEntity rowEntity(int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
                            if (row == null || cells == null || cells.isEmpty()) {
                                return null;
                            } else {
                                TestEntity entity = new TestEntity();
                                for (ExcelCell cell : cells) {
                                    switch (cell.getColIndex()) {
                                        case 0:
                                            entity.setColStr(cell.getStrValue());
                                            break;
                                        case 1:
                                            entity.setColInt(Integer.parseInt(cell.getStrValue()));
                                            break;
                                        case 2:
                                            entity.setColDbl(Double.parseDouble(cell.getStrValue()));
                                            break;
                                        case 3:
                                            entity.setColLng(Long.parseLong(cell.getStrValue()));
                                            break;
                                        case 4:
                                            entity.setColFlt(Float.parseFloat(cell.getStrValue()));
                                            break;
                                        case 5:
                                            entity.setColBol(Boolean.parseBoolean(cell.getStrValue()));
                                            break;
                                        case 6:
                                            entity.setColDat(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(cell.getStrValue()));
                                            break;
                                    }
                                }
                                return entity;
                            }
                        }

                        @Override
                        protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                            System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                            assert rowStart == 0;
                            assert rowEnd == 1;
                            assert list.size() == 1;
                            TestEntity entity = list.get(0);
                            System.out.println("解析到实体" + entity);
                            assert entity.getColStr().equals("str");
                            assert entity.getColInt() == 1;
                            assert entity.getColDbl() == 2;
                            assert entity.getColLng() == 3;
                            assert entity.getColFlt() == 4;
                            assert entity.isColBol();
                            assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity.getColDat()).equals("2019-01-01 00:00:00.000");
                        }
                    };
                } else {
                    return new SimpleSheetListener<TestEntity>(sheet) {
                        @Override
                        protected int titleRowCount() {
                            return 2;
                        }

                        @Override
                        protected boolean rowSkip(int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
                            return false;
                        }

                        @Override
                        protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                            System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                            assert rowStart == 2;
                            assert rowEnd == 4;
                            assert list.size() == 2;
                            System.out.println("解析到实体null");
                            assert list.get(0) == null;
                            TestEntity entity = list.get(1);
                            System.out.println("解析到实体" + entity);
                            assert entity.getColStr().equals("str1");
                            assert entity.getColInt() == 2;
                            assert entity.getColDbl() == 3;
                            assert entity.getColLng() == 4;
                            assert entity.getColFlt() == 5;
                            assert !entity.isColBol();
                            assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity.getColDat()).equals("2020-01-01 00:00:00.000");

                        }
                    };
                }
            }
        });
    }

    @Test
    public void testSimpleAdaptive() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleAdaptive.xls")), new SimpleExcelListener() {

            @Override
            public SimpleSheetListener<?> sheetListener(@Nonnull ExcelSheet sheet) {
                return new SimpleSheetListener<TestEntity>(sheet) {

                    @Override
                    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                        System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                        assert rowStart == 1;
                        assert rowEnd == 2;
                        assert list.size() == 1;
                        TestEntity entity = list.get(0);
                        System.out.println("解析到实体" + entity);
                        assert entity.getColStr().equals("str");
                        assert entity.getColInt() == 1;
                        assert entity.getColDbl() == 2;
                        assert entity.getColLng() == 3;
                        assert entity.getColFlt() == 4;
                        assert entity.isColBol();
                        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity.getColDat()).equals("2019-01-01 00:00:00.000");
                    }
                };
            }
        });
    }

    @Test(expected = RuntimeException.class)
    public void testSimpleThrow() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleMismatch.xls")), new SimpleExcelListener() {

            @Override
            public SimpleSheetListener<TestEntity> sheetListener(@Nonnull ExcelSheet sheet) {
                return new SimpleSheetListener<TestEntity>(sheet) {

                    @Override
                    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                    }
                };
            }
        });
    }

    @Test
    public void testSimpleDefault() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleMismatch.xls")), new SimpleExcelListener() {

            @Override
            public SimpleSheetListener<TestEntity> sheetListener(@Nonnull ExcelSheet sheet) {
                return new SimpleSheetListener<TestEntity>(sheet) {

                    @Nonnull
                    @Override
                    public Converter.MismatchPolicy mismatchPolicy() {
                        return Converter.MismatchPolicy.Default;
                    }

                    @Override
                    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
                        System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                        assert rowStart == 1;
                        assert rowEnd == 2;
                        assert list.size() == 1;
                        TestEntity entity = list.get(0);
                        System.out.println("解析到实体" + entity);
                        assert entity.getColStr().equals("nan");
                        assert entity.getColInt() == 0;
                        assert entity.getColDbl() == 0;
                        assert entity.getColLng() == 0;
                        assert entity.getColFlt() == 0;
                        assert !entity.isColBol();
                        assert entity.getColDat() == null;
                    }
                };
            }
        });
    }
}