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

public final class XlsxReader extends ExcelReader {

    private Builder builder;

    public XlsxReader(Builder builder) {
        this.builder = builder;
    }

    @Override
    public List<List<List<String>>> read(InputStream in) throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
        List<List<List<String>>> list = new ArrayList<>();
        List<String> sheets = builder.getSheets();
        int rowLength = builder.getRowLength();
        int fromRow = builder.getFromRow();
        if (fromRow < 0) {
            fromRow = 0;
        }
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
                list.add(readSheet(sheet, fromRow, rowLength));
            }
        } else {
            for (Sheet sheet : xssfWorkbook) {
                list.add(readSheet((XSSFSheet) sheet, fromRow, rowLength));
            }
        }
        return list;
    }

    private List<List<String>> readSheet(XSSFSheet sheet, int fromRow, int rowLength) {
        int fromColumn = builder.getFromColumn();
        int columnLength = builder.getColumnLength();
        if (fromColumn < 0) {
            fromColumn = 0;
        }
        List<List<String>> list = new ArrayList<>();
        int toRow;
        if (rowLength < 0) {
            toRow = sheet.getLastRowNum();
        } else {
            toRow = fromRow + rowLength;
        }
        for (int i = fromRow; i < toRow; i++) {
            XSSFRow row = sheet.getRow(i);
            list.add(readRow(row, fromColumn, columnLength));
        }
        return list;
    }

    private List<String> readRow(XSSFRow row, int fromColumn, int columnLength) {
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
        return list;
    }
}
