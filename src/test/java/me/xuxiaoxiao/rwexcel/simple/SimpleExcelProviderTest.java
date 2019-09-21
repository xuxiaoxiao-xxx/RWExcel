package me.xuxiaoxiao.rwexcel.simple;

import lombok.Data;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriterImpl;
import org.junit.Test;

import javax.annotation.Nonnull;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class SimpleExcelProviderTest {

    @Test
    public void sheetProvider() throws Exception {
        ExcelWriter excelWriter = new ExcelWriterImpl();
        excelWriter.write(new FileOutputStream("test.xls"), new SimpleExcelProvider() {
            @Nonnull
            @Override
            public ExcelWriter.Version version() {
                return ExcelWriter.Version.XLS;
            }


            @Override
            public SimpleSheetProvider<Entity> sheetProvider(ExcelSheet sheet) {

                return new SimpleSheetProvider<Entity>(sheet) {

                    @Override
                    public List<Entity> queryList(int lastRowIndex) {
                        List<Entity> entities = new ArrayList<>(100);
                        Entity entity1 = new Entity();
                        entity1.colA = "str";
                        entity1.colB = 1;
                        entity1.colC = 2.3;
                        entity1.colD = 4;
                        entity1.colE = 5.6f;
                        entity1.colF = true;
                        entity1.colG = new Date();
                        for (int i = 0; i < 100; i++) {
                            entities.add(entity1);
                        }
                        if (lastRowIndex < 60000) {
                            System.out.println("查询100条数据");
                            return entities;
                        } else {
                            return new LinkedList<>();
                        }
                    }
                };
            }
        });
    }

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
}