package me.xuxiaoxiao.rwexcel.simple;

import lombok.Data;

import java.util.Date;

@Data
public class TestEntity {
    @ExcelColumn("列str")
    private String colStr;

    @ExcelColumn("列int")
    private int colInt;

    @ExcelColumn("列dbl")
    private double colDbl;

    @ExcelColumn("列lng")
    private long colLng;

    @ExcelColumn("列flt")
    private float colFlt;

    @ExcelColumn("列bol")
    private boolean colBol;

    @ExcelColumn("列dat")
    private Date colDat;
}
