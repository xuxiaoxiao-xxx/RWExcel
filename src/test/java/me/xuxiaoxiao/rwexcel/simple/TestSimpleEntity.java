package me.xuxiaoxiao.rwexcel.simple;

import lombok.Data;

import java.util.Date;

@Data
public class TestSimpleEntity {
    @ExcelColumn("列str")
    private String colStr;

    @ExcelColumn("列int")
    private int colInt;

    @ExcelColumn("列long")
    private long colLong;

    @ExcelColumn("列float")
    private float colFloat;

    @ExcelColumn("列double")
    private double colDouble;

    @ExcelColumn("列bool")
    private boolean colBool;

    @ExcelColumn("列date")
    private Date colDate;
}
