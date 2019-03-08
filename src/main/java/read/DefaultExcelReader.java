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
import java.lang.reflect.Constructor;
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
        List<Sheet> sheets = sheets(workbook(in, builder.getType()));
        List<T> list = new ArrayList<>();
        for (Sheet sheet : sheets) {
            list.addAll(readSheet(sheet, function));
        }
        in.close();
        return list;
    }

    public <T> List<T> read(InputStream in, Class<T> clazz) throws Exception {
        Assert.notNull(in);
        Assert.notNull(clazz);
        if (!clazz.isAnnotationPresent(SimpleExcel.class)) {
            throw new ExcelRWException("Wrong class");
        }
        constructorValidate(clazz);
        List<Sheet> sheets = sheets(workbook(in, builder.getType()));
        List<T> list = new ArrayList<>();
        for (Sheet sheet : sheets) {
            list.addAll(readSheet(sheet, clazz));
        }
        in.close();
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

    private List<Sheet> sheets(Workbook workbook) {
        List<Sheet> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                list.add(workbook.getSheet(sheetName));
            }
        } else {
            for (Sheet sheet : workbook) {
                list.add(sheet);
            }
        }
        return list;
    }

    private <T> List<T> readSheet(Sheet sheet, Class<T> clazz) throws Exception {
        List<T> list = new ArrayList<>();
        toRow = getToRow(sheet);
        for (int i = builder.getFromRow(); i < toRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            list.add(readRow(row, clazz));
        }
        return list;
    }

    private <T> List<T> readSheet(Sheet sheet, Function<List<String>, T> function) {
        List<T> list = new ArrayList<>();
        toRow = getToRow(sheet);
        for (int i = builder.getFromRow(); i < toRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            list.add(readRow(row, function));
        }
        return list;
    }

    private <T> T readRow(Row row, Function<List<String>, T> function) {
        List<String> list = new ArrayList<>();
        toColumn = getToColumn(row);
        for (int i = builder.getFromColumn(); i < toColumn; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                list.add(builder.getCellFunction().apply(cell));
            }
        }
        return function.apply(list);
    }

    private <T> T readRow(Row row, Class<T> clazz) throws Exception {
        T target = clazz.newInstance();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            if (!field.isAnnotationPresent(Column.class)) {
                continue;
            }
            Column column = field.getAnnotation(Column.class);
            int index = column.index();
            Cell cell = row.getCell(index);
            Method writeMethod = new PropertyDescriptor(field.getName(), clazz).getWriteMethod();
            Class converter = column.converter();
            if (converter == Converter.class) {
                writeMethod.invoke(target, value(cell, field.getType()));
            } else {
                constructorValidate(converter);
                Object invoke = converter.getMethod(Converter.class.getMethods()[0].getName(), String.class).invoke(converter.newInstance(), value(cell, String.class));
                writeMethod.invoke(target, invoke);
            }
        }
        return target;
    }

    private int getToRow(Sheet sheet) {
        if (builder.getRowLength() < 0) {
            toRow = sheet.getLastRowNum() + 1;
        }
        return toRow;
    }

    private int getToColumn(Row row) {
        if (builder.getColumnLength() < 0) {
            toColumn = row.getLastCellNum() + 1;
        }
        return toColumn;
    }

    private <T> void constructorValidate(Class<T> clazz) {
        boolean flag = false;
        Constructor<?>[] constructors = clazz.getConstructors();
        for (Constructor<?> constructor : constructors) {
            if (constructor.getParameterCount() == 0) {
                flag = true;
            }
        }
        if (!flag) {
            throw new ExcelRWException("A parametric constructor method must be provided");
        }
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
