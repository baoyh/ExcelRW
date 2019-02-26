package read;

import annotation.Column;
import annotation.SimpleExcel;
import bean.Person;
import exception.ExcelRWException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class ReadPerson {
    public static void main(String[] args) throws IOException, IntrospectionException, InvocationTargetException, IllegalAccessException, InstantiationException, ExcelRWException {
        read(new FileInputStream("E:/test.xlsx"), Person.class);
    }

    public static <T> List<T> read(InputStream in, Class<T> clazz) throws IOException, IllegalAccessException, InstantiationException, IntrospectionException, InvocationTargetException, ExcelRWException {
        if (!clazz.isAnnotationPresent(SimpleExcel.class)) {
            throw new ExcelRWException("wrong class");
        }
        List<T> list = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        Workbook sheets = new XSSFWorkbook(in);
        for (Sheet sheet : sheets) {
            for (Row row : sheet) {
                T object = clazz.newInstance();
                for (Field field : fields) {
                    field.setAccessible(true);
                    Column column = field.getAnnotation(Column.class);
                    int index = column.index();
                    Cell cell = row.getCell(index);
                    PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
                    Method method = pd.getWriteMethod();
                    method.invoke(object, cell.toString());
                }
                list.add(object);
            }
        }
        System.out.println(list);
        return list;
    }
}
