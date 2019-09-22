package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelReaderImpl;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriterImpl;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class SimpleExcelProviderTest {
    @Test
    public void testSimpleSheets() throws Exception {
        ByteArrayOutputStream baOutStream = new ByteArrayOutputStream();

        ExcelWriter excelWriter = new ExcelWriterImpl();
        excelWriter.write(baOutStream, new SimpleExcelProvider() {
            @Nonnull
            @Override
            public ExcelWriter.Version version() {
                return ExcelWriter.Version.XLS;
            }

            @Nonnull
            @Override
            public ExcelSheet[] sheets() {
                return new ExcelSheet[]{new ExcelSheet(0, "Sheet1"), new ExcelSheet(1, "Sheet2")};
            }

            @Override
            public SimpleSheetProvider<TestEntity> sheetProvider(ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    return new SimpleSheetProvider<TestEntity>(sheet) {

                        @Nullable
                        @Override
                        public List<TestEntity> queryList(int lastRowIndex) {
                            if (lastRowIndex < titleRowCount()) {
                                List<TestEntity> entities = new ArrayList<>(100);
                                TestEntity entity1 = new TestEntity();
                                entity1.setColStr("str");
                                entity1.setColInt(1);
                                entity1.setColDbl(2);
                                entity1.setColLng(3);
                                entity1.setColFlt(4);
                                entity1.setColBol(true);
                                try {
                                    entity1.setColDat(new SimpleDateFormat("yyyy-MM-dd").parse("2019-01-01"));
                                } catch (ParseException e) {
                                    e.printStackTrace();
                                }
                                entities.add(entity1);
                                return entities;
                            } else {
                                return null;
                            }
                        }
                    };
                } else {
                    return new SimpleSheetProvider<TestEntity>(sheet) {

                        @Override
                        protected boolean entitySkip(int lastRowIndex, @Nullable TestEntity entity) {
                            return false;
                        }

                        @Nullable
                        @Override
                        public List<TestEntity> queryList(int lastRowIndex) {
                            if (lastRowIndex < titleRowCount()) {
                                List<TestEntity> entities = new ArrayList<>(100);
                                entities.add(null);
                                TestEntity entity1 = new TestEntity();
                                entity1.setColStr("str1");
                                entity1.setColInt(2);
                                entity1.setColDbl(3);
                                entity1.setColLng(4);
                                entity1.setColFlt(5);
                                entity1.setColBol(false);
                                try {
                                    entity1.setColDat(new SimpleDateFormat("yyyy-MM-dd").parse("2020-01-01"));
                                } catch (ParseException e) {
                                    e.printStackTrace();
                                }
                                entities.add(entity1);
                                return entities;
                            } else {
                                return null;
                            }
                        }
                    };
                }
            }
        });

        ExcelReader excelReader = new ExcelReaderImpl();
        excelReader.read(new ByteArrayInputStream(baOutStream.toByteArray()), new SimpleExcelListenerTest.TestSheetsListener());
    }
}