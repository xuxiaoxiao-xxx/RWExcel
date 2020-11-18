package me.xuxiaoxiao.rwexcel;

import lombok.Data;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.simple.ExcelColumn;
import me.xuxiaoxiao.rwexcel.simple.SimpleSheetListener;
import org.junit.Test;

import javax.annotation.Nonnull;
import java.util.Date;
import java.util.List;
import java.util.Objects;

public class DemoTest {

    @Data
    public static class DemoUser {
        @ExcelColumn("姓名")
        private String name;
        @ExcelColumn("年龄")
        private String age;
        @ExcelColumn("出生日期")
        private Date birthday;
    }

    @Test
    public void demo() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("demo.xls")), new SimpleSheetListener<DemoUser>() {

            @Override
            protected void onList(int rowStart, int rowEnd, @Nonnull List<DemoUser> list) {
                System.out.println("解析XLS成列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                for (DemoUser demoUser : list) {
                    System.out.println("解析到用户：" + demoUser);
                }
            }
        });
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("demo.xlsx")), new SimpleSheetListener<DemoUser>() {

            @Override
            protected void onList(int rowStart, int rowEnd, @Nonnull List<DemoUser> list) {
                System.out.println("解析XLSX成列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
                for (DemoUser demoUser : list) {
                    System.out.println("解析到用户：" + demoUser);
                }
            }
        });
    }
}
