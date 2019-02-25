package read;

import exception.ExcelRWException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public final class XlsxReader extends ExcelReader {

    private int toRow;
    private int toColumn;
    private Builder builder;

    public XlsxReader(Builder builder) {
        toRow = builder.getFromRow() + builder.getRowLength();
        toColumn = builder.getFromColumn() + builder.getColumnLength();
        this.builder = builder;
    }

    @Override
    public <R> List<R> read(InputStream in, Function<List<String>, R> function) throws IOException, ExcelRWException {
        Assert.notNull(function);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
        List<R> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
                list.addAll(readSheet(sheet, function));
            }
        } else {
            for (Sheet sheet : xssfWorkbook) {
                list.addAll(readSheet((XSSFSheet) sheet, function));
            }
        }
        return list;
    }

    private <R> List<R> readSheet(XSSFSheet sheet, Function<List<String>, R> function) {
        List<R> list = new ArrayList<>();
        if (builder.getRowLength() < 0) {
            toRow = sheet.getLastRowNum() + 1;
        }
        for (int i = builder.getFromRow(); i < toRow; i++) {
            XSSFRow row = sheet.getRow(i);
            list.add(readRow(row, function));
        }
        return list;
    }

    private <R> R readRow(XSSFRow row, Function<List<String>, R> function) {
        if (builder.getColumnLength() < 0) {
            toColumn = row.getLastCellNum() + 1;
        }
        List<String> list = new ArrayList<>();
        for (int i = builder.getFromColumn(); i < toColumn; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                list.add(builder.getCellFunction().apply(cell));
            }
        }
        return function.apply(list);
    }
}
