package me.xuxiaoxiao.rwexcel.simple;

import lombok.Getter;
import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;

/**
 * 简单Sheet数据源，将List转换成sheet数据
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetProvider<T> implements ExcelWriter.Provider {
    private final List<T> list = new LinkedList<>();

    @Getter
    private final ExcelSheet sheet;
    @Getter
    private final Class<T> clazz;

    protected final int colFirst, colLast;
    protected final Map<Field, Integer> fMapper = new HashMap<>();
    protected final Map<Field, Converter> cMapper = new HashMap<>();

    /**
     * 创建一个sheet数据源
     *
     * @param sheet 当前sheet信息
     */
    public SimpleSheetProvider(ExcelSheet sheet) {
        this.sheet = sheet;
        this.clazz = entityClass();

        int colF = Integer.MAX_VALUE, colL = Integer.MIN_VALUE;
        boolean adaptive = true;
        List<Field> fieldList = new LinkedList<>();
        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);
                    fieldList.add(field);

                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        adaptive = false;
                        if (excelColumn.index() < colF) {
                            colF = excelColumn.index();
                        }
                        if (colL < excelColumn.index()) {
                            colL = excelColumn.index();
                        }
                        fMapper.put(field, excelColumn.index());
                    } else if (!adaptive) {
                        throw new IllegalArgumentException(String.format("ExcelColumn的列 %s 未设置序号", excelColumn.value()));
                    }
                }
            }
        }
        if (adaptive) {
            Collections.sort(fieldList, new Comparator<Field>() {
                @Override
                public int compare(Field o1, Field o2) {
                    Class<?> c1 = o1.getDeclaringClass();
                    Class<?> c2 = o2.getDeclaringClass();
                    if (c1.equals(c2)) {
                        return 0;
                    } else if (c1.isAssignableFrom(c2)) {
                        return -1;
                    } else {
                        return 1;
                    }
                }
            });
            colF = 0;
            colL = fieldList.size() - 1;
            for (int i = 0; i < fieldList.size(); i++) {
                fMapper.put(fieldList.get(i), i);
            }
        }
        this.colFirst = colF;
        this.colLast = colL;
        if (fMapper.isEmpty()) {
            throw new IllegalArgumentException("SimpleSheetProvider的模板参数中未找到任何使用@Excelolumn标记的属性");
        }
    }

    @Nonnull
    @Override
    public ExcelWriter.Version version() {
        return ExcelWriter.Version.XLSX;
    }

    @Nullable
    @Override
    public ExcelSheet provideSheet(int lastSheetIndex) {
        return null;
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
        if (lastRowIndex + 1 < titleRowCount()) {
            return titleRow(lastRowIndex);
        } else {
            if (list.isEmpty()) {
                List<T> temp = queryList(lastRowIndex);
                if (temp != null) {
                    this.list.addAll(temp);
                }
            }
            while (list.size() > 0 && entitySkip(lastRowIndex, list.get(0))) {
                list.remove(0);
            }
            if (list.isEmpty()) {
                return null;
            } else {
                return entityRow(lastRowIndex, list.get(0));
            }
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
        if (row.getRowIndex() < titleRowCount()) {
            return titleCells(row);
        } else {
            try {
                return entityCells(row, list.remove(0));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 获取实体类的class对象
     *
     * @return 实体类的class对象
     */
    @Nonnull
    protected Class<T> entityClass() {
        if (this.clazz == null) {
            Class<?> providerClass = this.getClass();
            if (providerClass.getGenericSuperclass() instanceof ParameterizedType) {
                Type[] typeArguments = ((ParameterizedType) providerClass.getGenericSuperclass()).getActualTypeArguments();
                if (typeArguments != null && typeArguments.length > 0) {
                    Type paramType = typeArguments[0];
                    if (paramType instanceof Class) {
                        return (Class<T>) paramType;
                    }
                }
            }
            throw new IllegalStateException("未能自动识别SimpleSheetProvider的模板参数");
        } else {
            return this.clazz;
        }
    }

    /**
     * 转换异常的处理方法
     *
     * @return 转换异常的处理方法
     */
    @Nonnull
    public Converter.MismatchPolicy mismatchPolicy() {
        return Converter.MismatchPolicy.Throw;
    }

    /**
     * 标题行数量
     *
     * @return 标题行数量
     */
    protected int titleRowCount() {
        return 1;
    }

    /**
     * 标题行信息
     *
     * @param lastRowIndex 上一个行号
     * @return 标题行信息
     */
    @Nullable
    protected ExcelRow titleRow(int lastRowIndex) {
        ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
        excelRow.setColFirst(this.colFirst);
        excelRow.setColLast(this.colLast);
        return excelRow;
    }

    /**
     * 标题行的单元格列表
     *
     * @param row 标题行信息
     * @return 标题行的单元格列表
     */
    @Nonnull
    protected List<ExcelCell> titleCells(@Nonnull ExcelRow row) {
        List<ExcelCell> cells = new LinkedList<>();
        for (Map.Entry<Field, Integer> entry : this.fMapper.entrySet()) {
            cells.add(new ExcelCell(this.sheet.getShtIndex(), row.getRowIndex(), entry.getValue(), entry.getKey().getAnnotation(ExcelColumn.class).value()));
        }
        Collections.sort(cells, new Comparator<ExcelCell>() {
            @Override
            public int compare(ExcelCell o1, ExcelCell o2) {
                return o1.getColIndex() - o2.getColIndex();
            }
        });
        return cells;
    }

    /**
     * 是否要跳过某个实体
     *
     * @param lastRowIndex 上一个行号
     * @param entity       实体信息
     * @return 是否要跳过
     */
    protected boolean entitySkip(int lastRowIndex, @Nullable T entity) {
        //跳过null实体
        return entity == null;
    }

    /**
     * 获取实体对应的行信息
     *
     * @param lastRowIndex 上一个行号
     * @param entity       实体信息
     * @return 实体对应的行信息
     */
    @Nullable
    protected ExcelRow entityRow(int lastRowIndex, @Nullable T entity) {
        ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
        excelRow.setColFirst(this.colFirst);
        excelRow.setColLast(this.colLast);
        return excelRow;
    }

    /**
     * 获取实体对应的单元格列表
     *
     * @param row    实体对应的行信息
     * @param entity 实体信息
     * @return 实体对应的单元格列表
     * @throws Exception 转换时可能发生异常
     */
    @Nonnull
    protected List<ExcelCell> entityCells(@Nonnull ExcelRow row, @Nullable T entity) throws Exception {
        if (entity == null) {
            return new LinkedList<>();
        } else {
            List<ExcelCell> cells = new LinkedList<>();
            for (Map.Entry<Field, Integer> entry : this.fMapper.entrySet()) {
                Field field = entry.getKey();
                Converter converter = cMapper.get(field);
                if (converter == null) {
                    converter = field.getAnnotation(ExcelColumn.class).converter().newInstance();
                    cMapper.put(field, converter);
                }
                try {
                    cells.add(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), fMapper.get(field), converter.obj2str(field, field.get(entity))));
                } catch (Exception e) {
                    if (mismatchPolicy() == Converter.MismatchPolicy.Throw) {
                        throw e;
                    }
                }
            }
            Collections.sort(cells, new Comparator<ExcelCell>() {
                @Override
                public int compare(ExcelCell o1, ExcelCell o2) {
                    return o1.getColIndex() - o2.getColIndex();
                }
            });
            return cells;
        }
    }

    /**
     * 查询实体列表，进行一个批次的处理，会调用多次
     *
     * @param lastRowIndex 最后一行的行号，首次为-1
     * @return 一批次的实体列表
     */
    public abstract List<T> queryList(int lastRowIndex);
}
