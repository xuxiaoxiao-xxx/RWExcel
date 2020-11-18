package me.xuxiaoxiao.rwexcel.simple;

import javax.annotation.Nonnull;
import java.text.SimpleDateFormat;
import java.util.List;

public class TestSimpleListener extends SimpleSheetListener<TestSimpleEntity> {
    @Override
    protected void onList(int rowStart, int rowEnd, @Nonnull List<TestSimpleEntity> list) {
        System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
        assert rowStart == 1;
        assert rowEnd == 8;
        assert list.size() == 6;
        TestSimpleEntity entity1 = list.get(0);
        System.out.println("解析到实体" + entity1);
        assert entity1.getColStr().equals("str1");
        assert entity1.getColDouble() == 1;
        assert entity1.getColFloat() == 2;
        assert entity1.getColLong() == 3;
        assert entity1.getColInt() == 4;
        assert entity1.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity1.getColDate()).equals("2019-01-01 00:00:00.000");

        TestSimpleEntity entity2 = list.get(1);
        System.out.println("解析到实体" + entity2);
        assert entity2.getColStr().equals("str2");
        assert entity2.getColDouble() == 3.4;
        assert entity2.getColFloat() == 2.3f;
        assert entity2.getColLong() == 31312312;
        assert entity2.getColInt() == 22;
        assert entity2.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity2.getColDate()).equals("2019-01-01 00:00:22.000");

        TestSimpleEntity entity3 = list.get(2);
        System.out.println("解析到实体" + entity3);
        assert entity3.getColStr().equals("str3");
        assert entity3.getColDouble() == 22.3;
        assert entity3.getColFloat() == 0;
        assert entity3.getColLong() == 0;
        assert entity3.getColInt() == 0;
        assert entity3.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity3.getColDate()).equals("2000-01-01 00:00:00.000");

        TestSimpleEntity entity4 = list.get(3);
        System.out.println("解析到实体" + entity4);
        assert entity4.getColStr().equals("str4");
        assert entity4.getColDouble() == 11.2;
        assert entity4.getColFloat() == 0;
        assert entity4.getColLong() == 0;
        assert entity4.getColInt() == 0;
        assert !entity4.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity4.getColDate()).equals("2001-01-01 00:00:00.000");

        TestSimpleEntity entity5 = list.get(4);
        System.out.println("解析到实体" + entity5);
        assert entity5.getColStr().equals("str5");
        assert entity5.getColDouble() == 0;
        assert entity5.getColFloat() == 0;
        assert entity5.getColLong() == 12312;
        assert entity5.getColInt() == 0;
        assert !entity5.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity5.getColDate()).equals("2002-01-02 00:00:00.000");

        TestSimpleEntity entity6 = list.get(5);
        System.out.println("解析到实体" + entity6);
        assert entity6.getColStr() == null;
        assert entity6.getColDouble() == 0;
        assert entity6.getColFloat() == 0;
        assert entity6.getColLong() == 0;
        assert entity6.getColInt() == 11;
        assert !entity6.isColBool();
        assert new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS").format(entity6.getColDate()).equals("2003-01-02 03:00:00.000");
    }
}