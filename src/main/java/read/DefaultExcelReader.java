package read;

import annotation.Column;
import annotation.SimpleExcel;
import convert.BooleanConverter;
import convert.Converter;
import convert.NumberConverter;
import convert.StringConverter;
import exception.ExcelRWException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import type.Excel;
import util.Assert;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public final class DefaultExcelReader extends ExcelReader {

    private int toRow;
    private int toColumn;
    private Builder builder;

    public DefaultExcelReader(Builder builder) {
        toRow = builder.getFromRow() + builder.getRowLength();
        toColumn = builder.getFromColumn() + builder.getColumnLength();
        if (builder.getCellFunction() == null) {
            builder.cellFunction(cell -> {return (String) value(cell, String.class);});
        }
        this.builder = builder;
    }

    @Override
    public <T> List<T> read(InputStream in, Function<List<String>, T> function) throws IOException, ExcelRWException {
        Assert.notNull(function);
        Assert.notNull(in);
        Workbook workbook = workbook(in, builder.getType());
        List<T> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                Sheet sheet = workbook.getSheet(sheetName);
                list.addAll(readSheet(sheet, function));
            }
        } else {
            for (Sheet sheet : workbook) {
                list.addAll(readSheet(sheet, function));
            }
        }
        return list;
    }

    @Override
    public <T> List<T> read(InputStream in, Class<T> clazz) throws Exception {
        if (!clazz.isAnnotationPresent(SimpleExcel.class)) {
            throw new ExcelRWException("Wrong class");
        }
        List<T> list = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        Workbook workbook = workbook(in, builder.getType());
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                T object = clazz.newInstance();
                for (Field field : fields) {
                    field.setAccessible(true);
                    Class<?> fieldType = field.getType();
                    if (!field.isAnnotationPresent(Column.class)) {
                        continue;
                    }
                    Column column = field.getAnnotation(Column.class);
                    int index = column.index();
                    Cell cell = row.getCell(index);
                    PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
                    Method method = pd.getWriteMethod();
                    Class converter = column.converter();
                    if (converter == Converter.class) {
                        method.invoke(object, value(cell, fieldType));
                    } else {
                        Object o = converter.newInstance();
                        Method method1 = converter.getMethod(Converter.class.getMethods()[0].getName(), String.class);
                        Object invoke = method1.invoke(o, value(cell, String.class));
                        method.invoke(object, invoke);
                    }
                }
                list.add(object);
            }
        }
        return list;
    }

    private Workbook workbook(InputStream in, Excel type) throws IOException {
        Workbook workbook = null;
        switch (type) {
            case XLS:
                workbook = new HSSFWorkbook(in);
                break;
            case XLSX:
                workbook = new XSSFWorkbook(in);
                break;
        }
        return workbook;
    }

    private <R> List<R> readSheet(Sheet sheet, Function<List<String>, R> function) {
        List<R> list = new ArrayList<>();
        if (builder.getRowLength() < 0) {
            toRow = sheet.getLastRowNum() + 1;
        }
        for (int i = builder.getFromRow(); i < toRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            list.add(readRow(row, function));
        }
        return list;
    }

    private <R> R readRow(Row row, Function<List<String>, R> function) {
        if (builder.getColumnLength() < 0) {
            toColumn = row.getLastCellNum() + 1;
        }
        List<String> list = new ArrayList<>();
        for (int i = builder.getFromColumn(); i < toColumn; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                list.add(builder.getCellFunction().apply(cell));
            }
        }
        return function.apply(list);
    }

    private Object value(Cell cell, Class<?> type) {
        CellType typeEnum = cell.getCellTypeEnum();
        switch (typeEnum) {
            case NUMERIC:
                return new NumberConverter<>(type).convert(cell.getNumericCellValue());
            case STRING:
                return new StringConverter<>(type).convert(cell.getStringCellValue());
            case BOOLEAN:
                return new BooleanConverter<>(type).convert(cell.getBooleanCellValue());
            case FORMULA:
                return new NumberConverter<>(type).convert(cell.getNumericCellValue());
            case BLANK:
                return null;
            case _NONE:
                throw new ExcelRWException("Excel type is none");
            case ERROR:
                throw new ExcelRWException("Excel type is error");
        }
        return null;
    }
}
