package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelReaderImpl;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Objects;

public class SimpleExcelListenerTest {

    @Test
    public void testSimpleSheets() throws Exception {
        ExcelReader reader = new ExcelReaderImpl();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheets.xls")), new SimpleExcelListener() {

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
        });
    }

    @Test
    public void testSimpleTitle() throws Exception {
        ExcelReader reader = new ExcelReaderImpl();
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
                        protected Field entityField(int rowIndex, int colIndex) {
                            try {
                                switch (colIndex) {
                                    case 0:
                                        return entityClass().getDeclaredField("colStr");
                                    case 1:
                                        return entityClass().getDeclaredField("colInt");
                                    case 2:
                                        return entityClass().getDeclaredField("colDbl");
                                    case 3:
                                        return entityClass().getDeclaredField("colLng");
                                    case 4:
                                        return entityClass().getDeclaredField("colFlt");
                                    case 5:
                                        return entityClass().getDeclaredField("colBol");
                                    case 6:
                                        return entityClass().getDeclaredField("colDat");
                                }
                            } catch (NoSuchFieldException e) {
                                e.printStackTrace();
                            }
                            throw new RuntimeException("未找到Field");
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
        ExcelReader reader = new ExcelReaderImpl();
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
        ExcelReader reader = new ExcelReaderImpl();
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
        ExcelReader reader = new ExcelReaderImpl();
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