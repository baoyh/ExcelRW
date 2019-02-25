package read;

import exception.ExcelRWException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import type.Excel;
import util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public final class DefaultExcelReader extends ExcelReader {

    private int toRow;
    private int toColumn;
    private Builder builder;

    public DefaultExcelReader(Builder builder) {
        toRow = builder.getFromRow() + builder.getRowLength();
        toColumn = builder.getFromColumn() + builder.getColumnLength();
        this.builder = builder;
    }

    @Override
    public <R> List<R> read(InputStream in, Excel type, Function<List<String>, R> function) throws IOException, ExcelRWException {
        Assert.notNull(function);
        Assert.notNull(in);
        Workbook workbook = workbook(in, type);
        List<R> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                Sheet sheet = workbook.getSheet(sheetName);
                list.addAll(readSheet(sheet, function));
            }
        } else {
            for (Sheet sheet : workbook) {
                list.addAll(readSheet(sheet, function));
            }
        }
        return list;
    }

    private Workbook workbook(InputStream in, Excel type) throws IOException {
        Workbook workbook = null;
        switch (type) {
            case XLS:
                workbook = new HSSFWorkbook(in);
                break;
            case XLSX:
                workbook = new XSSFWorkbook(in);
                break;
        }
        return workbook;
    }

    private <R> List<R> readSheet(Sheet sheet, Function<List<String>, R> function) {
        List<R> list = new ArrayList<>();
        if (builder.getRowLength() < 0) {
            toRow = sheet.getLastRowNum() + 1;
        }
        for (int i = builder.getFromRow(); i < toRow; i++) {
            list.add(readRow(sheet.getRow(i), function));
        }
        return list;
    }

    private <R> R readRow(Row row, Function<List<String>, R> function) {
        if (builder.getColumnLength() < 0) {
            toColumn = row.getLastCellNum() + 1;
        }
        List<String> list = new ArrayList<>();
        for (int i = builder.getFromColumn(); i < toColumn; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                list.add(builder.getCellFunction().apply(cell));
            }
        }
        return function.apply(list);
    }
}
