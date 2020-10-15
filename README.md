# RWExcel
高性能Excel流式解析/导出框架，完全屏蔽03、07版本Excel差异。

# 特性描述
* 完全屏蔽03、07版本Excel差异，自动识别03、07版本Excel
* 自动寻列：根据Excel标题行名称和Java属性注解的名称，自动将Excel某列对应到Java类的某属性
* 解析和导出时均将单元格内容视作纯文本
* 支持流式和非流式方式解析和导出
* 暂不支持单元格样式和批注等的解析和导出

|         |  03版本（xls）  |  07版本（xlsx）  |
|  :---:  |  :---:  | :---:  |
|  解析性能  |  高(流式)  |  高(流式)  |
|  导出性能  |  低(伪流式)  |  高(流式)  |

# 最新版本
```xml
<dependency>
    <groupId>me.xuxiaoxiao</groupId>
    <artifactId>rwexcel</artifactId>
    <version>1.1.0</version>
</dependency>
```

# 使用示例
### Excel示例
```
标题行    |列str	|列int	|列dbl	|列lng	|列flt	|列bol	|列dat
数据行    |str	|1	|2	|3	|4	|TRUE	|2019-01-01 00:00:00
```
### Java类示例
```java
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
```
### SimpleSheetListener
SimpleSheetListener将一个Excel的所有Sheet按照相同的规则解析，示例代码如下
```java
@Test
public void testSimpleSheetListener() throws Exception {
    ExcelReader reader = new ExcelStreamReader();
    reader.read(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleAdaptive.xls"), new SimpleSheetListener<TestEntity>() {

        @Override
        protected void onList(int rowStart, int rowEnd, @Nonnull List<TestEntity> list) {
            System.out.println("解析到列表：rowStart=" + rowStart + "，rowEnd=" + rowEnd);
            for (TestEntity entity : list) {
                System.out.println("解析到实体：" + entity);
            }
        }
    });
}
```
运行结果
```
解析到列表：rowStart=1，rowEnd=2
解析到实体：TestEntity(colStr=str, colInt=1, colDbl=2.0, colLng=3, colFlt=4.0, colBol=true, colDat=Tue Jan 01 00:00:00 CST 2019)
```
### FAQ
##### Q:Excel中每个Sheet需要不同的解析规则怎么办
A:请使用SimpleExcelListener

##### Q:Sheet中列数量和含义不固定怎么办
A:请自行实现ExcelReader.Listener接口

##### Q:我想把Excel某列的枚举值，如：男、女，映射成Java类的整数属性，如：1，2
A:请继承并重写Converter类，并在注解属性时指定Converter