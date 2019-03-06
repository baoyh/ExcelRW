package convert;

public class MyConverter implements ExcelConverter<String> {

    @Override
    public String convert(String s) {
        return s + "success";
    }
}
