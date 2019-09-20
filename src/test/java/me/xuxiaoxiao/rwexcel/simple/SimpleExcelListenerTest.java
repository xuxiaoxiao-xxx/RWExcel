package me.xuxiaoxiao.rwexcel.simple;

import lombok.Data;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelReaderImpl;
import org.junit.Test;

import java.util.Date;
import java.util.List;
import java.util.Objects;

public class SimpleExcelListenerTest {

    @Data
    public static class Entity {
        @ExcelColumn("列A")
        private String colA;

        @ExcelColumn("列B")
        private int colB;

        @ExcelColumn("列C")
        private double colC;

        @ExcelColumn("列D")
        private long colD;

        @ExcelColumn("列E")
        private float colE;

        @ExcelColumn("列F")
        private boolean colF;

        @ExcelColumn("列G")
        private Date colG;
    }

    @Test
    public void sheetListener() throws Exception {
        ExcelReader reader = new ExcelReaderImpl();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xls")), new SimpleExcelListener() {

            @Override
            public SimpleSheetListener<?> sheetListener(ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    return new SimpleSheetListener<Entity>(sheet) {
                        @Override
                        protected void onList(int shtIndex, int rowStart, int rowEnd, List<Entity> list) {
                            for (Entity entity : list) {
                                System.out.println("解析到实体" + entity);
                            }
                        }
                    };
                } else {
                    return null;
                }
            }
        });
    }
}