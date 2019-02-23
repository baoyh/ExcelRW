package read;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public final class XlsxReader extends ExcelReader {

    private Builder builder;

    public XlsxReader(Builder builder) {
        this.builder = builder;
    }

    @Override
    public <R> List<List<R>> read(InputStream in, Function<List<String>, R> function) throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
        List<List<R>> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        int rowLength = builder.getRowLength();
        int fromRow = builder.getFromRow();
        if (fromRow < 0) {
            fromRow = 0;
        }
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
                list.add(readSheet(sheet, fromRow, rowLength, function));
            }
        } else {
            for (Sheet sheet : xssfWorkbook) {
                list.add(readSheet((XSSFSheet) sheet, fromRow, rowLength, function));
            }
        }
        return list;
    }

    private <R> List<R> readSheet(XSSFSheet sheet, int fromRow, int rowLength, Function<List<String>, R> function) {
        int fromColumn = builder.getFromColumn();
        int columnLength = builder.getColumnLength();
        if (fromColumn < 0) {
            fromColumn = 0;
        }
        List<R> list = new ArrayList<>();
        int toRow;
        if (rowLength < 0) {
            toRow = sheet.getLastRowNum();
        } else {
            toRow = fromRow + rowLength;
        }
        for (int i = fromRow; i < toRow; i++) {
            XSSFRow row = sheet.getRow(i);
            list.add(readRow(row, fromColumn, columnLength, function));
        }
        return list;
    }

    private <R> R readRow(XSSFRow row, int fromColumn, int columnLength, Function<List<String>, R> function) {
        int toColumn;
        if (columnLength < 0) {
            toColumn = row.getLastCellNum();
        } else {
            toColumn = fromColumn + columnLength;
        }
        List<String> list = new ArrayList<>();
        for (int i = fromColumn; i < toColumn; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell != null) {
                list.add(cell.toString());
            }
        }
        return function.apply(list);
    }
}
