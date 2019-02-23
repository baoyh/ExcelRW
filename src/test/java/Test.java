import exception.ExcelRWException;
import read.ExcelReader;
import read.XlsxReader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class Test {
    public static void main(String[] args) throws ExcelRWException, IOException {
        ExcelReader.Builder build = new ExcelReader.Builder().fromRow(1).rowLength(2).build();
        XlsxReader xlsxReader = new XlsxReader(build);
        List<List<List<String>>> read = xlsxReader.read(new FileInputStream("C:/Work/test.xlsx"));
        System.out.println(read);
    }
}
