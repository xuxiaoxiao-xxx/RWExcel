# RWExcel
高性能Excel流式解析/导出框架，完全屏蔽03、07版本Excel差异。

# 特性描述
|  特性  |  03版本（xls）  |  07版本（xlsx）  |
| :---:|  :---:  | :---:  |
|  解析性能  |  高(流式)  |  高(流式)  |
|  导出性能  |  低(伪流式)  |  高(流式)  |
|  单元格内容  |  解析成字符串/写出字符串  |  解析成字符串/写出字符串  |
|  单元格样式  |  暂不支持解析/导出  |  暂不支持解析/导出  |
|  单元格批注  |  不支持解析/导出  |  不支持解析/导出  |

# 即将开发
* UserModel（非流式解析/导出）支持
* 单元格样式支持

# 最新版本
```xml
<dependency>
    <groupId>me.xuxiaoxiao</groupId>
    <artifactId>rwexcel</artifactId>
    <version>1.0.0</version>
</dependency>
```

# 使用示例
#### SimpleExcelListener
SimpleExcelListener 配合 SimpleSheetListener 使用，能够将Excel的Sheet按行解析成实体列表。
当然，需要在实体类中使用@ExcelColumn注解标注属性。并且该监听器默认会认为Sheet的第一行为标题行。
1. **自动寻列模式**：如果实体中的所有@ExcelColumn注解都没有指定列序号，那么监听器将会根据Sheet第一行和列标题进行自动匹配。
2. 重写SimpleSheetListener中的方法能够修改**标题行解析、是否跳过某行、行解析为实体**等方法的默认实现。
3. **默认的Converter**：@ExcelColumn中可以指定某个属性的解析和导出方式，默认的Converter可以支持int,double,String,Date等大部分情况。

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


public class SimpleExcelListenerTest {
    
    @Test
    public void demo() throws Exception {
        new ExcelStreamReader().read(new FileInputStream("testSimpleSheets.xls"), new SimpleExcelListener() {

            @Nullable
            @Override
            public SimpleSheetListener<TestEntity> sheetListener(@Nonnull ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    //解析第一个Sheet，每个Sheet需要一个SimpleSheetListener
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
}
```
* SimpleExcelListener输出
```
解析到列表：rowStart=1，rowEnd=2
TestEntity(colStr=str, colInt=1, colDbl=2.0, colLng=3, colFlt=4.0, colBol=true, colDat=Tue Jan 01 00:00:00 CST 2019)
```

#### ExcelReader.Listener
ExcelReader.Listener 是ExcelReader的监听器接口，如果SimpleExcelListener无法满足你的需求，你可以实现自己的解析监听器

```java
public class ListenerTest {

    @Test
    public void demo() throws Exception {
        new ExcelStreamReader().read(new FileInputStream("test.xls"), new ExcelReader.Listener() {

            @Override
            public void onSheetStart(@Nonnull ExcelSheet sheet) {
                System.out.println("开始解析Sheet：" + sheet.getShtName());
            }

            @Override
            public void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
                System.out.println("解析row：rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
                for (ExcelCell cell : cells) {
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                }
            }

            @Override
            public void onSheetEnd(@Nonnull ExcelSheet sheet) {
                System.out.println("结束解析Sheet：" + sheet.getShtName());
            }
        });
    }
}
```
* ExcelReader.Listener输出
```
开始解析Sheet：Sheet1
解析row：rowIndex=0，colFirst=0，colLast=0
colIndex=0，strValue=1A
解析row：rowIndex=1，colFirst=0，colLast=1
colIndex=0，strValue=2A
colIndex=1，strValue=2B
解析row：rowIndex=2，colFirst=0，colLast=2
colIndex=0，strValue=3A
colIndex=1，strValue=3B
colIndex=2，strValue=3C
解析row：rowIndex=3，colFirst=1，colLast=2
colIndex=1，strValue=4B
colIndex=2，strValue=4C
解析row：rowIndex=4，colFirst=2，colLast=2
colIndex=2，strValue=5C
结束解析Sheet：Sheet1
```