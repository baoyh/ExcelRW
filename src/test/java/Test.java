import exception.ExcelRWException;
import read.ExcelReader;
import read.DefaultExcelReader;
import type.Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.function.Function;

public class Test {
    public static void main(String[] args) throws ExcelRWException, IOException {

        Function<List<String>, Person> asPerson = list -> {
            Person p = new Person();
            p.name = list.get(0);
            p.gender = list.get(1);
            p.address = list.get(2);
            p.phone = list.get(3);
            return p;
        };

        ExcelReader.Builder build = new ExcelReader.Builder().fromRow(1).rowLength(2).fromColumn(1).columnLength(2).build();
        ExcelReader reader = new DefaultExcelReader(build);
        List<List<String>> read1 = reader.read(new FileInputStream("C:/Work/test.xls"), Excel.XLS);
        System.out.println(read1);
    }

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
}
