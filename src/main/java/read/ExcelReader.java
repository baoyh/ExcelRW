package read;

import exception.ExcelRWException;
import org.apache.poi.ss.usermodel.Cell;
import type.Excel;
import util.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.function.Function;

public abstract class ExcelReader {

    private Function<List<String>, List<String>> DEFAULT_ROW_FUNCTION = Function.identity();

    public static class Builder {
        private List<String> sheets;
        private int fromRow = 0;
        private int rowLength = -1;
        private int fromColumn = 0;
        private int columnLength = -1;
        private Excel type = Excel.XLSX;

        private Function<Cell, String> cellFunction;

        public Builder cellFunction(Function<Cell, String> cellFunction) throws ExcelRWException {
            Assert.notNull(cellFunction);
            this.cellFunction = cellFunction;
            return this;
        }

        public Builder sheets(List<String> names) throws ExcelRWException {
            Assert.notNull(names);
            this.sheets = names;
            return this;
        }

        public Builder fromRow(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromRow = i;
            return this;
        }

        public Builder rowLength(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.rowLength = i;
            return this;
        }

        public Builder fromColumn(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.fromColumn = i;
            return this;
        }

        public Builder columnLength(int i) throws ExcelRWException {
            Assert.greaterThanZero(i);
            this.columnLength = i;
            return this;
        }

        public Builder type(Excel type) throws ExcelRWException {
            Assert.notNull(type);
            this.type = type;
            return this;
        }

        public Builder build() {
            return this;
        }

        public int getRowLength() {
            return rowLength;
        }

        public int getColumnLength() {
            return columnLength;
        }

        public List<String> getSheets() {
            return sheets;
        }

        public int getFromRow() {
            return fromRow;
        }

        public int getFromColumn() {
            return fromColumn;
        }

        public Function<Cell, String> getCellFunction() {
            return cellFunction;
        }

        public Excel getType() {
            return type;
        }
    }

    public abstract <T> List<T> read(InputStream in, Function<List<String>, T> rowFunction) throws IOException, ExcelRWException;

    public List<List<String>> read(InputStream in) throws IOException, ExcelRWException { return read(in, DEFAULT_ROW_FUNCTION); }

    public abstract <T> List<T> read(InputStream in, Class<T> clazz) throws Exception;

}
