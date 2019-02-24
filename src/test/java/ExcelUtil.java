import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUtil {

    static class Person {
        String name;
        String phone;
        String address;
        String gender;


        @Override
        public String toString() {
            return "Person{" +
                    "name='" + name + '\'' +
                    ", phone='" + phone + '\'' +
                    ", address='" + address + '\'' +
                    ", gender='" + gender + '\'' +
                    '}';
        }
    }

    public static void main(String[] args) throws Exception {
        String path = "C:/Work/test.xlsx";
        //test1(path);

        Function<List<String>, List<String>> asList = Function.identity();

        Function<List<String>, Map<String, String>> asMap = strings -> {
            Map<String, String> map = new HashMap<>();
            for (String string : strings) {
                map.put(string, string);
            }
            return map;
        };

        Function<List<String>, Person> asPerson = list -> {
            Person p = new Person();
            p.name = list.get(0);
            p.gender = list.get(1);
            p.address = list.get(2);
            p.phone = list.get(3);
            return p;
        };

        Function<Cell, String> toString = cell -> {
            switch (cell.getCellTypeEnum()) {
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                default:
                    return cell.getStringCellValue();
            }
        };

        readAsStream(path).map(row -> {
            List<Cell> list = new ArrayList<>();
            row.cellIterator().forEachRemaining(list::add);
            return list;
        }).map(cells -> {
            List<String> list = new ArrayList<>();
            cells.iterator().forEachRemaining(cell -> list.add(toString.apply(cell)));//这里加上不同类型的处理加到Object中
            return list;
        })
                //.map(asList)
                .map(asMap)
                //.map(asPerson)
                .forEach(System.out::println);


    }

    private static void test1(String path) throws Exception {

        List<List<String>> list1 = readXlsx(path, Function.identity());
        list1.stream().forEach(System.out::println);


        System.out.println("-----------------------");
        List<Map<String, String>> list2 = readXlsx(path, strings -> {
            Map<String, String> map = new HashMap<>();
            for (String string : strings) {
                map.put(string, string);
            }
            return map;
        });

        list2.stream().forEach(System.out::println);
        System.out.println("-----------------------");
        List<Person> people = readXlsx(path, list -> {
            Person p = new Person();
            p.name = list.get(0);
            p.gender = list.get(1);
            p.address = list.get(2);
            p.phone = list.get(3);
            return p;
        });

        people.stream().forEach(System.out::println);
    }


    public static <R> List<R> readXlsx(String path, Function<List<String>, R> function) throws Exception {
        FileInputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        List<R> result = new ArrayList<>();
        for (Sheet xssfSheet : xssfWorkbook) {
            if (xssfSheet != null) {
                XSSFRow xssfRow;
                short minColIx = 0;
                short maxColIx = 0;
                for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    xssfRow = ((XSSFSheet) xssfSheet).getRow(rowNum);
                    if (xssfRow == null) {
                        continue;
                    }
                    try {
                        minColIx = xssfRow.getFirstCellNum();
                        maxColIx = xssfRow.getLastCellNum();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    List<String> rowList = new ArrayList<>();
                    XSSFCell cell;
                    String value;
                    for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                        cell = xssfRow.getCell(colIx);
                        if (cell != null) {
                            value = cell.toString();
                            if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                                Long longVal = Math.round(cell.getNumericCellValue());
                                double doubleVal = cell.getNumericCellValue();
                                if (Double.parseDouble(longVal + ".0") == doubleVal) {   //判断是否含有小数位.0
                                    value = String.valueOf(longVal);
                                }
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                                    SimpleDateFormat sdf = null;
                                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
                                            .getBuiltinFormat("h:mm")) {
                                        sdf = new SimpleDateFormat("HH:mm");
                                    } else {// 日期
                                        sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                                    }
                                    Date date = cell.getDateCellValue();
                                    value = sdf.format(date);
                                }
                            }

                            rowList.add(value);
                        } else {
                            rowList.add("");
                        }
                    }
                    result.add(function.apply(rowList));
                }
            }

        }
        is.close();
        return result;
    }


    public static Stream<Row> readAsStream(String path) throws Exception {
        FileInputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(
                sheet.iterator(), Spliterator.ORDERED | Spliterator.NONNULL), false);
    }
} 
