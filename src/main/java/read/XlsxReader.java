package read;

import exception.ExcelRWException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public final class XlsxReader extends ExcelReader {

    @Override
    public List<List<List<String>>> read(InputStream in, Builder builder) throws ExcelRWException, IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
        List<String> sheets = builder.getSheets();
        if (sheets != null && !sheets.isEmpty()) {
            for (String sheetName : sheets) {
                XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
            }
        }
        return null;
    }

    private void readSheet(XSSFSheet sheet, Builder builder) {
        int toRow = builder.getToRow();
        int fromRow = builder.getFromRow();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = toRow; i < fromRow; i++) {
            XSSFRow row = sheet.getRow(i);
        }
    }

    private void readRow(XSSFRow row, Builder builder) {
        short fromColumn = builder.getFromColumn();
        short toColumn = builder.getToColumn();
        short lastCellNum = row.getLastCellNum();
        for (short i = 0; i < lastCellNum; i++) {

        }
    }
}
