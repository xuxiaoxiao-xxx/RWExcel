package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeMap;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/19 16:41]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetListener<T> implements ExcelReader.Listener {
    private final ExcelSheet sheet;
    private final int cache;
    private final List<T> list;
    private final Class<T> clazz;
    private final boolean adaptive;

    private int rowStart = -1, rowEnd = -1;

    protected final Converter converter = new Converter();
    protected final TreeMap<Integer, Field> fMapper = new TreeMap<>();
    protected final TreeMap<Integer, Converter> cMapper = new TreeMap<>();


    public SimpleSheetListener(ExcelSheet sheet) {
        this(sheet, 100);
    }

    public SimpleSheetListener(ExcelSheet sheet, int cache) {
        this.sheet = sheet;
        this.cache = cache;
        this.list = new ArrayList<>(cache);
        this.clazz = entityClass();

        boolean adaptive = true;
        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);

                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        adaptive = false;
                        if (fMapper.containsKey(excelColumn.index())) {
                            throw new IllegalArgumentException(String.format("ExcelColumn列冲突，有多个序号为 %d 的列", excelColumn.index()));
                        } else {
                            fMapper.put(excelColumn.index(), field);
                        }
                    } else if (!adaptive) {
                        throw new IllegalArgumentException(String.format("ExcelColumn的列 %s 未设置序号", excelColumn.value()));
                    }
                }
            }
        }
        this.adaptive = adaptive;
    }

    @Override
    public final void onSheetStart(@Nonnull ExcelSheet sheet) {
    }

    @Override
    public final void onRow(@Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        int shtIndex = sheet.getShtIndex();
        while (rowEnd < row.getRowIndex() - 1) {
            try {
                //处理空白行
                if (!rowSkip(shtIndex, rowEnd, null, null)) {
                    list.add(rowEntity(shtIndex, rowEnd, null, null));
                    if (list.size() >= cache) {
                        flush();
                    }
                }
                rowEnd++;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }

        try {
            if (!rowSkip(shtIndex, row.getRowIndex(), row, cells)) {
                list.add(rowEntity(shtIndex, row.getRowIndex(), row, cells));
                if (list.size() >= cache) {
                    flush();
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        flush();
    }

    private void flush() {
        onList(sheet.getShtIndex(), rowStart, rowEnd, list);
        rowStart = rowEnd;
        list.clear();
    }

    protected Class<T> entityClass() {
        if (this.clazz == null) {
            Class<?> listenerClass = this.getClass();
            if (listenerClass.getGenericSuperclass() instanceof ParameterizedType) {
                Type[] typeArguments = ((ParameterizedType) listenerClass.getGenericSuperclass()).getActualTypeArguments();
                if (typeArguments != null && typeArguments.length > 0) {
                    Type paramType = typeArguments[0];
                    if (paramType instanceof Class) {
                        return (Class<T>) paramType;
                    }
                }
            }
            throw new IllegalStateException("未能自动识别SimpleExcelListener的模板参数");
        } else {
            return this.clazz;
        }
    }

    @Nullable
    protected Field entityField(int rowIndex, int colIndex) {
        return fMapper.get(colIndex);
    }

    protected boolean rowSkip(int shtIndex, int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
        if (rowIndex == 0 && adaptive && row != null && cells != null && !cells.isEmpty()) {
            //处理第一行（标题行），如果是自动寻列模式，要找一下属性和列的对应关系
            for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
                Field[] fields = i.getDeclaredFields();
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        field.setAccessible(true);

                        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                        for (ExcelCell cell : cells) {
                            if (excelColumn.value().equals(cell.getStrValue())) {
                                if (fMapper.containsKey(cell.getColIndex())) {
                                    throw new IllegalArgumentException(String.format("ExcelColumn自动寻列冲突，有多个成员域title为 %s", excelColumn.value()));
                                } else {
                                    fMapper.put(cell.getColIndex(), field);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
        return rowIndex == 0 || row == null;
    }

    protected T rowEntity(int shtIndex, int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
        if (row == null || cells == null || cells.isEmpty()) {
            return null;
        } else {
            T entity = this.clazz.newInstance();
            for (ExcelCell cell : cells) {
                Field field = entityField(cell.getRowIndex(), cell.getColIndex());
                if (field != null) {
                    Class<? extends Converter> cClass = field.getAnnotation(ExcelColumn.class).converter();
                    if (Converter.class.equals(cClass)) {
                        field.set(entity, converter.str2obj(field, cell.getStrValue()));
                    } else {
                        Converter converter = cMapper.get(cell.getColIndex());
                        if (converter == null) {
                            converter = cClass.newInstance();
                            cMapper.put(cell.getColIndex(), converter);
                        }
                        field.set(entity, converter.str2obj(field, cell.getStrValue()));
                    }
                }
            }
            return entity;

        }
    }

    protected abstract void onList(int shtIndex, int rowStart, int rowEnd, List<T> list);
}
