import bean.Person;
import read.ExcelReader;
import read.DefaultExcelReader;
import type.Excel;

import java.io.FileInputStream;
import java.util.List;

public class Test {
    public static void main(String[] args) throws Exception {
        String path = "E:/test.xlsx";
        ExcelReader.Builder build = new ExcelReader.Builder().type(Excel.XLSX).fromRow(1).rowLength(2).build();
        ExcelReader reader = new DefaultExcelReader(build);
        List<Person> read = reader.read(new FileInputStream(path), Person.class);
        //List<List<String>> read = reader.read(new FileInputStream(path));
        System.out.println(read);
    }

}




