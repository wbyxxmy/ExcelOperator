package com.yexj.excelOperator.utils;

import com.yexj.excelOperator.annotation.Column;

import java.lang.reflect.Field;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;

public final class ColumnUtils {

    /**
     * client should make sure clazz is not null and is annotated by Column
     */
    public static Column annotationObject(Class<?> clazz) {
        Objects.requireNonNull(clazz);

        Column anno = clazz.getAnnotation(Column.class);
        Objects.requireNonNull(anno);

        return anno;
    }

    /**
     * 按照定义/priority顺序返回Column -> Field
     *
     */
    public static LinkedHashMap<Column, Field> findOrderedAnnotatedColumns(Class<?> clazz, int export) {
        Map<Integer, LinkedHashMap<Column, Field>> columnsByPriority = new TreeMap<Integer, LinkedHashMap<Column, Field>>();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            Column column = field.getAnnotation(Column.class);
            if (column != null && (column.export() == export || column.export() == 0)) {
                LinkedHashMap<Column, Field> columns = columnsByPriority.get(column.priority());
                if (columns == null) {
                    columns = new LinkedHashMap<Column, Field>();
                    columnsByPriority.put(column.priority(), columns);
                }

                columns.put(column, field);
            }
        }

        LinkedHashMap<Column, Field> columns = new LinkedHashMap<Column, Field>();
        for (LinkedHashMap<Column, Field> cs : columnsByPriority.values()) {
            columns.putAll(cs);
        }

        return columns;
    }
}
