/*
import annotation.Column;
import annotation.SimpleExcel;
import bean.Person;
import convert.*;
import exception.ExcelRWException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.PropertyDescriptor;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class ReadPerson {
    public static void main(String[] args) throws Exception {
        read(new FileInputStream("E:/test.xlsx"), Person.class);
    }

    public static <T> List<T> read(InputStream in, Class<T> clazz) throws Exception {
        if (!clazz.isAnnotationPresent(SimpleExcel.class)) {
            throw new ExcelRWException("Wrong class");
        }
        List<T> list = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        Workbook sheets = new XSSFWorkbook(in);
        for (Sheet sheet : sheets) {
            for (Row row : sheet) {
                T object = clazz.newInstance();
                for (Field field : fields) {
                    field.setAccessible(true);
                    Class<?> type = field.getType();
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
                        method.invoke(object, f(type, cell));
                    } else {
                        Object o = converter.newInstance();
                        Method method1 = converter.getMethod(Converter.class.getMethods()[0].getName(), String.class);
                        Object invoke = method1.invoke(o, f(String.class, cell));
                        method.invoke(object, invoke);
                    }
                }
                list.add(object);
            }
        }
        System.out.println(list);
        return list;
    }

    private static Object f(Class<?> type, Cell cell) {
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
*/
