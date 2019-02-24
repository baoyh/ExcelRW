import exception.ExcelRWException;
import read.ExcelReader;
import read.XlsxReader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.function.Function;

public class Test {
    public static void main(String[] args) throws ExcelRWException, IOException {
        ExcelReader.Builder build = new ExcelReader.Builder().fromRow(0).build();
        ExcelReader xlsxReader = new XlsxReader(build);
        Function<List<String>, Person> asPerson = list -> {
            Person p = new Person();
            p.name = list.get(0);
            p.gender = list.get(1);
            p.address = list.get(2);
            p.phone = list.get(3);
            return p;
        };
        List<Person> read = xlsxReader.read(new FileInputStream("C:/Work/test.xlsx"), asPerson);
        System.out.println(read);
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
