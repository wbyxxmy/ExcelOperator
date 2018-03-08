package com.yexj.excelOperator.utils;

import com.yexj.excelOperator.annotation.*;
import com.yexj.excelOperator.exception.IllegalAnnotationException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

/**
 *
 * Created by xinjian.ye on 2018/3/7.
 */
public class WorkbookUtils {
    private WorkbookUtils(){

    }

    public static SXSSFWorkbook createWorkbook() {
        return new SXSSFWorkbook();
    }

    public static <T>
    SXSSFWorkbook createWorkbook(Class<T> clazz, T ... sources) throws IllegalAnnotationException {
        return createWorkbook(clazz, Arrays.asList(sources));
    }

    public static <T>
    SXSSFWorkbook createWorkbook(Class<T> clazz, Collection<T> sources) throws IllegalAnnotationException {
        SXSSFWorkbook workbook = createWorkbook();
        addNewSheet(workbook, clazz, sources);

        return workbook;
    }

    public static org.apache.poi.ss.usermodel.Sheet addNewSheet(SXSSFWorkbook workbook, String name) {
        Objects.requireNonNull(workbook);
        Objects.requireNonNull(name);

        return workbook.createSheet(name);
    }

    public static <T>
    org.apache.poi.ss.usermodel.Sheet addNewSheet(SXSSFWorkbook workbook,
                                                  Class<T> clazz,
                                                  Collection<T> sources) throws IllegalAnnotationException {
        return addNewSheet(workbook, clazz, sources, true);
    }

    public static <T>
    org.apache.poi.ss.usermodel.Sheet addNewSheet(SXSSFWorkbook workbook,
                                                  Class<T> clazz,
                                                  Collection<T> sources,
                                                  boolean activeNewSheet) throws IllegalAnnotationException {
        Objects.requireNonNull(workbook);
        Objects.requireNonNull(clazz);

        com.yexj.excelOperator.annotation.Sheet sheetAnno = SheetUtils.annotationObject(clazz);

        org.apache.poi.ss.usermodel.Sheet sheet = addNewSheet(workbook, sheetAnno.name());
        if (activeNewSheet) {
            workbook.setActiveSheet(workbook.getSheetIndex(sheet));
        }

        LinkedHashMap<Column, Field> annotatedColumns = ColumnUtils.findOrderedAnnotatedColumns(clazz);
        if (! annotatedColumns.isEmpty()) {
            int rowIndex = 0;
            if (sheetAnno.containsHeader()) {
                CellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                createHeader(sheet, rowIndex++, annotatedColumns.keySet(), sheetAnno.containsSequence(),style);
            }

            if (sources != null && ! sources.isEmpty()) {
                Collection<Field> annotatedFields = annotatedColumns.values();
                int sourceSequence = 1;
                for (T src : sources) {
                    if (src == null) {
                        continue;
                    }

                    createRow(sheet, rowIndex++, src, annotatedFields, sheetAnno.containsSequence(),
                            String.valueOf(sourceSequence++));
                }
            }
        }

        return sheet;
    }

    public static Row appendRow(org.apache.poi.ss.usermodel.Sheet sheet) {
        Objects.requireNonNull(sheet);
        return sheet.createRow(sheet.getLastRowNum() + 1);
    }

    public static void appendRow(org.apache.poi.ss.usermodel.Sheet sheet,
                                 Object source,
                                 String sourceId) throws IllegalAnnotationException {
        Objects.requireNonNull(sheet);
        Objects.requireNonNull(source);

        Class<?> clazz = source.getClass();
        LinkedHashMap<Column, Field> annotatedColumns = ColumnUtils.findOrderedAnnotatedColumns(clazz);
        if (! annotatedColumns.isEmpty()) {
            int rowIndex = sheet.getLastRowNum() + 1;
            boolean containsSequence = SheetUtils.annotationObject(clazz).containsSequence();
            createRow(sheet, rowIndex, source, annotatedColumns.values(), containsSequence, sourceId);
        }
    }

    private static void createHeader(org.apache.poi.ss.usermodel.Sheet sheet,
                                     int rowIndex,
                                     Collection<Column> columns,
                                     boolean containsSequence, CellStyle cellStyle) {
        Row row = sheet.createRow(rowIndex);

        int columnIndex = 0;
        if (containsSequence) {
            CellUtil.createCell(row, columnIndex++, "").setCellStyle(cellStyle);
        }

        for (Column column : columns) {
            CellUtil.createCell(row, columnIndex++, column.name()).setCellStyle(cellStyle);
        }
    }

    private static void createRow(org.apache.poi.ss.usermodel.Sheet sheet,
                                  int rowIndex,
                                  Object source,
                                  Collection<Field> fields,
                                  boolean containsSequence,
                                  String sourceId) throws IllegalAnnotationException {
        Row row = sheet.createRow(rowIndex);

        int columnIndex = 0;
        if (containsSequence) {
            row.createCell(columnIndex++).setCellValue(sourceId != null ? sourceId : "");
        }

        for (Field field : fields) {
            setupCell(row.createCell(columnIndex++), source, field);
        }
    }

    private static void setupCell(Cell cell, Object object, Field field) throws IllegalAnnotationException {
        boolean accessible = field.isAccessible();

        try {
            field.setAccessible(true);

            Object v = field.get(object);
            if (v == null) {
                // DO NOTHING for null
            }
            else if (v instanceof String) {
                cell.setCellValue((String) v);
            }
            else if (v instanceof Number) {
                cell.setCellValue(((Number) v).doubleValue());
            }
            else if (v instanceof Boolean) {
                cell.setCellValue((Boolean) v);
            }
            else if (v instanceof Date) {
                cell.setCellValue((Date) v);
            }
            else if (v instanceof Calendar) {
                cell.setCellValue((Calendar) v);
            }
            else {
                throw new IllegalAnnotationException("Unsupported column type of field: " + field);
            }
        } catch (Exception e) {
            // NOTE:
            // since setAccessible(true) is invoked before get field's value, should NEVER visit here
            throw new RuntimeException("Unable to access field: " + field);
        } finally {
            field.setAccessible(accessible);
        }
    }


    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    public static String getStringCellValue(Cell cell) {
        String strCell = "";
        if(cell == null) return null;
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                double v = cell.getNumericCellValue();
                BigDecimal dou = new BigDecimal(v);
                BigDecimal str = new BigDecimal(v+"");
                strCell = str.compareTo(dou) == 0 ? dou.toString() : str.toString();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("")) {
            return "";
        }
        return strCell;
    }
}
